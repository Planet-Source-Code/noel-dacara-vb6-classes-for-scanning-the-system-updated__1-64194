VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScanRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --------------------------------------------------------------------------------
' Name:     cScanRegistry class
' Author:   Noel A. Dacara (noeldacara@hotmail.com)
' Date:     22 August 2005
'
' You can freely use and distribute this file with or without permission provided
' that credits herein are kept unmodified and having acknowledgements is observed.
'
' Caution:  Incorrectly editing the registry may severely damage your system.
'           Before using this module, backup any valued data on your computer.
'
' Notes:    Filter feature for keys is applicable for first level keys only which
'           is after the assigned start path. Preceding subkeys are automatically
'           included or scanned if permitted by the ScanSubkeys property.
'
'           Filters could also contain wildcards and can be as many as you want.
'           Just separate filter patters with the pipe character. E.g. Run*|E?x
'
'           If FullRegistryScan is set, StartPath and CustomScan are ignored.
'
' Tips:     To attain the maximum scanning speed of your machine with this class,
'           you may need to close unnecessary programs running in your computer..
'
' Inquiries or comments regarding my works and whatsoever are greatly appreciated.
' --------------------------------------------------------------------------------

' API Declarations
Private Declare Function GetQueueStatus Lib "user32.dll" (ByVal fuFlags As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long

' Modified API Declarations
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByVal lpReserved As Long, ByRef lpcSubKeys As Long, ByRef lpcbMaxSubKeyLen As Long, ByRef lpcbMaxClassLen As Long, ByRef lpcValues As Long, ByRef lpcbMaxValueNameLen As Long, ByRef lpcbMaxValueLen As Long, ByRef lpcbSecurityDescriptor As Long, ByRef lpftLastWriteTime As Any) As Long

' API Constants
Private Const ERROR_SUCCESS             As Long = 0&
Private Const KEY_CREATE_LINK           As Long = &H20
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_SET_VALUE             As Long = &H2
Private Const MAX_PATH                  As Long = 260
Private Const REG_OPTION_NON_VOLATILE   As Long = 0
Private Const REG_OPTION_RESERVED       As Long = 0
Private Const STANDARD_RIGHTS_ALL       As Long = &H1F0000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const KEY_ALL_ACCESS            As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

' GetQueueStatus Flag
Private Const QS_ALLEVENTS As Long = &HBF

' User-Defined API Enum
Public Enum ECustomRegistryPaths
    SCAN_ADDREMOVELISTS = 0     ' Add/Remove Programs
    SCAN_CUSTOMCONTROLS = 1     ' Custom Controls
    SCAN_FILEEXTENSIONS = 2     ' File Extensions
    SCAN_HELPRESOURCES = 3      ' Help and Resources
    SCAN_SHAREDDLLS = 4         ' Shared DLLs
    SCAN_SHELLFOLDERS = 5       ' Registry Integrity
    SCAN_SOFTWAREPATHS = 6      ' Software Locations
    SCAN_STARTUPKEYS = 7        ' Startup Programs
    SCAN_WINDOWSFONTS = 8       ' Windows Fonts
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private SCAN_ADDREMOVELISTS, SCAN_CUSTOMCONTROLS, SCAN_FILEEXTENSIONS
    Private SCAN_HELPRESOURCES, SCAN_SHAREDDLLS, SCAN_SHELLFOLDERS
    Private SCAN_SOFTWAREPATHS, SCAN_STARTUPKEYS, SCAN_WINDOWSFONTS
#End If

Public Enum EScanRegistryRoots
    SCAN_CLASSES_ROOT = &H80000000  ' These 4 registry roots would be enough...
    SCAN_CURRENT_USER = &H80000001
    SCAN_LOCAL_MACHINE = &H80000002
    SCAN_USERS = &H80000003
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private SCAN_CLASSES_ROOT, SCAN_CURRENT_USER, SCAN_LOCAL_MACHINE, SCAN_USERS
#End If

' Variable Declarations
Private m_bCancel       As Boolean
Private m_bDataDeleted  As Boolean
Private m_bCustomScan   As Boolean
Private m_bFilterData   As Boolean
Private m_bFilterKeys   As Boolean
Private m_bFullScan     As Boolean
Private m_bKeyDeleted   As Boolean
Private m_bPaused       As Boolean
Private m_bScanning     As Boolean
Private m_bScanSubKeys  As Boolean
Private m_eClassRoot    As EScanRegistryRoots
Private m_eCustomPath   As ECustomRegistryPaths
Private m_lFilterData   As Long
Private m_lFilterKeys   As Long
Private m_lScanDeep     As Long
Private m_lStartDeep    As Long
Private m_lTotalData    As Long
Private m_lTotalKeys    As Long
Private m_sFilterData() As String
Private m_sFilterKeys() As String
Private m_sScanPath     As String

Public Event BuildingDataList(Index As Long, Total As Long)
Attribute BuildingDataList.VB_Description = "Occurs while class is building data list but before sending data to events."
'Occurs while class is building data list but before sending data to events.
Public Event BuildingKeyList(Index As Long, Total As Long)
Attribute BuildingKeyList.VB_Description = "Occurs while class is building key list but before sending keys to events."
'Occurs while class is building key list but before sending keys to events.
Public Event CurrentData(Value As String, Key As String, Root As EScanRegistryRoots, Delete As Boolean)
Attribute CurrentData.VB_Description = "Occurs when a registry data is scanned or processed."
'Occurs when a registry data is scanned or processed.
Public Event CurrentKey(Key As String, Root As EScanRegistryRoots, Delete As Boolean)
Attribute CurrentKey.VB_Description = "Occurs when a registry key has been scanned or processed for data."
'Occurs when a registry key has been scanned or processed for data.
Public Event DoneScanning(TotalData As Long, TotalKeys As Long)
Attribute DoneScanning.VB_Description = "Occurs when scanning is done either successfully or cancelled."
'Occurs when scanning is done either successfully or cancelled.

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private Index, Total, Value, Key, Delete, Root
    Private DataCount, KeyCount, TotalData, TotalKeys
#End If

' //-- Properties --//

Public Property Get ClassRoot() As EScanRegistryRoots
Attribute ClassRoot.VB_Description = "Returns/sets a value to determine the registry key root type to scan."
'Returns/sets a value to determine the registry key root type to scan.
    ClassRoot = m_eClassRoot
End Property

Public Property Let ClassRoot(Value As EScanRegistryRoots)
    m_eClassRoot = Value
End Property

Public Property Get CustomScan() As Boolean
Attribute CustomScan.VB_Description = "Returns/sets a value to determine where to scan for keys and values on the predefined registry locations."
'Returns/sets a value to determine where to scan for keys and values on the predefined registry locations.
    CustomScan = m_bCustomScan
End Property

Public Property Let CustomScan(Value As Boolean)
    m_bCustomScan = Value
End Property

Public Property Get CustomScanPath() As ECustomRegistryPaths
Attribute CustomScanPath.VB_Description = "Returns/sets whether to scan custom registry keys for data."
'Returns/sets whether to scan custom registry keys for data.
    CustomScanPath = m_eCustomPath
End Property

Public Property Let CustomScanPath(Value As ECustomRegistryPaths)
    m_eCustomPath = Value
End Property

Public Property Get DataDeleted() As Boolean
Attribute DataDeleted.VB_Description = "Returns True if the previous data with Delete argument set to True was successfully removed from registry."
'Returns True if the previous data with Delete argument set to True was successfully removed from registry.
    DataDeleted = m_bDataDeleted
End Property

Public Property Get FilterData() As String
Attribute FilterData.VB_Description = "Returns/sets the filter pattern to scan for specific data."
'Returns/sets the filter pattern to scan for specific data.
    FilterData = Join(m_sFilterData, "|")
End Property

Public Property Let FilterData(Value As String)
    m_sFilterData = Split(UCase$(Value), "|")
    m_lFilterData = UBound(m_sFilterData)
    m_bFilterData = (m_lFilterData <> -1)
End Property

Public Property Get FilterKeys() As String
Attribute FilterKeys.VB_Description = "Returns/sets the filter pattern to scan for specific keys."
'Returns/sets the filter pattern to scan for specific keys.
    FilterKeys = Join(m_sFilterKeys, "|")
End Property

Public Property Let FilterKeys(Value As String)
    m_sFilterKeys = Split(UCase$(Value), "|")
    m_lFilterKeys = UBound(m_sFilterKeys)
    m_bFilterKeys = (m_lFilterKeys <> -1)
End Property

Public Property Get FullRegistryScan() As Boolean
Attribute FullRegistryScan.VB_Description = "Returns/sets a value to determine whether to scan the whole registry for data."
'Returns/sets a value to determine whether to scan the whole registry for data.
    FullRegistryScan = m_bFullScan
End Property

Public Property Let FullRegistryScan(Value As Boolean)
    m_bFullScan = Value
End Property

Public Property Get KeyDeleted() As Boolean
Attribute KeyDeleted.VB_Description = "Returns True if the previous key with Delete argument set to True was successfully removed from registry."
'Returns True if the previous key with Delete argument set to True was successfully removed from registry.
    KeyDeleted = m_bKeyDeleted
End Property

Public Property Get Paused() As Boolean
Attribute Paused.VB_Description = "Returns True if scanning process has been temporarily stopped."
'Returns True if scanning process has been temporarily stopped.
    Paused = m_bPaused
End Property

Public Property Get ScanPath() As String
Attribute ScanPath.VB_Description = "Returns/sets the registry key where to start scanning."
'Returns/sets the registry key where to start scanning.
    ScanPath = m_sScanPath
End Property

Public Property Let ScanPath(Value As String)
    m_sScanPath = Value
End Property

Public Property Get ScanDeep() As Long
Attribute ScanDeep.VB_Description = "Returns/sets a value to determine how deep from the startpath to scan."
'Returns/sets a value to determine how deep from the startpath to scan.
    ScanDeep = m_lScanDeep
End Property

Public Property Let ScanDeep(Value As Long)
    If Value < 0 Then ' Restrict users from setting invalid values
        Value = 0
    End If
    m_lScanDeep = Value
End Property

Public Property Get Scanning() As Boolean
Attribute Scanning.VB_Description = "Returns True if class is currently on scanning state."
'Returns True if class is currently on scanning state.
    Scanning = m_bScanning
End Property

Public Property Get ScanSubKeys() As Boolean
Attribute ScanSubKeys.VB_Description = "Returns/sets whether to include subkeys when scanning."
'Returns/sets whether to include subkeys when scanning.
    ScanSubKeys = m_bScanSubKeys
End Property

Public Property Let ScanSubKeys(Value As Boolean)
    m_bScanSubKeys = Value
End Property

Public Property Get TotalData() As String
Attribute TotalData.VB_Description = "Returns the current number of data that has been scanned."
'Returns the current number of data that has been scanned.
    TotalData = m_lTotalData
End Property

Public Property Get TotalKeys() As Long
Attribute TotalKeys.VB_Description = "Returns the current number of keys that has been scanned."
'Returns the current number of keys that has been scanned.
    TotalKeys = m_lTotalKeys
End Property

Public Sub BeginScanning(Optional ClassRoot, Optional ScanPath, Optional CustomScanPath)
Attribute BeginScanning.VB_Description = "Start scanning for keys and values on the specified registry key."
'Start scanning for keys and values on the specified registry key.
    If (m_bPaused) Then
        ResumeScanning ' Resume scanning instead
        Exit Sub
    End If
    
    If (m_bScanning) Then
        Exit Sub ' Avoid cascading scans
    End If
    
    m_lTotalData = 0
    m_lTotalKeys = 0
    
    If (m_bCustomScan Or m_bFullScan) Then
        If (Not IsMissing(CustomScanPath)) Then
            m_eCustomPath = CustomScanPath
        End If
        
        Dim p_sFilterKeys   As String
        Dim p_bScanSubKeys  As Boolean
        
        p_sFilterKeys = FilterKeys
        p_bScanSubKeys = m_bScanSubKeys
        m_bScanSubKeys = True
        
        InitializeCustomScanning
        
        FilterKeys = p_sFilterKeys
        m_bScanSubKeys = p_bScanSubKeys
    Else
        If (Not IsMissing(ClassRoot)) Then
            m_eClassRoot = ClassRoot
        End If
        
        If (Not IsMissing(ScanPath)) Then
            m_sScanPath = ScanPath
        End If
        
        ValidateKey m_sScanPath
        InitializeScanning
        GetData m_sScanPath
        
        If (m_bScanSubKeys) Then
            GetKeys m_sScanPath
        End If
    End If
    
    FinalizeScanning
End Sub

Public Sub CancelScanning()
Attribute CancelScanning.VB_Description = "Abort the whole scanning process that has been started."
'Abort the whole scanning process that has been started.
    m_bCancel = True
    ResumeScanning ' Resume if scanning has been paused
End Sub

Public Sub PauseScanning()
Attribute PauseScanning.VB_Description = "Pause scanning process."
'Pause scanning process.
    If (m_bScanning) Then
        m_bPaused = True
    End If
End Sub

Public Sub ResumeScanning()
Attribute ResumeScanning.VB_Description = "Resume paused scanning process."
'Resume paused scanning process.
    If (m_bPaused) Then
        m_bPaused = False
    End If
End Sub

Public Sub ValidateKey(Key As String)
Attribute ValidateKey.VB_Description = "Validate a registry key from common typographical mistakes."
'Validate a registry key from common typographical mistakes.
    Key = Replace$(Key, "/", "\")
    
    While InStr(1, Key, "\\")
        Key = Replace$(Key, "\\", "\")
    Wend
    
    If (Key Like "*\") Then
        Key = Left$(Key, Len(Key) - 1)
    End If
    
    If (Key Like "\*") Then
        Key = Mid$(Key, 2)
    End If
End Sub

Public Function DeleteValue( _
        Root As EScanRegistryRoots, _
        Path As String, _
        Value As String) As Boolean
Attribute DeleteValue.VB_Description = "Delete a registry data specified. Returns True if successful."
'Delete a registry data specified. Returns True if successful.
    Dim p_hKey  As Long
    Dim p_lRet  As Long
    
    p_lRet = RegOpenKeyEx(Root, Path, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_hKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        p_lRet = RegDeleteValue(p_hKey, Value)
        RegCloseKey p_hKey
    End If
    
    DeleteValue = (p_lRet = ERROR_SUCCESS)
End Function

Public Function DeleteKey(Root As EScanRegistryRoots, Path As String) As Boolean
Attribute DeleteKey.VB_Description = "Attempt to delete registry key specified. Returns True if successful."
'Delete a registry key specified. Returns True if successful.
'Key must have no subkeys inorder for it to be deleted.
    Dim p_hKey          As Long
    Dim p_hParentKey    As String
    Dim p_lRet          As Long
    
    p_lRet = InStrRev(Path, "\")
    
    If (p_lRet > 0) Then
        p_hParentKey = Left$(Path, p_lRet - 1)
    End If
    
    p_lRet = RegOpenKeyEx(Root, p_hParentKey, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_hKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        p_hParentKey = Right$(Path, Len(Path) - Len(p_hParentKey) - 1)
        p_lRet = RegDeleteKey(p_hKey, p_hParentKey)
        
        RegCloseKey p_hKey
    End If
    
    DeleteKey = (p_lRet = ERROR_SUCCESS)
End Function

' //-- Private Procedures --//

Private Sub InitializeCustomScanning()
    Dim p_eClassRoots() As EScanRegistryRoots
    Dim p_lCtr          As Long
    Dim p_lScanDeeps()  As Long
    Dim p_sScanPaths()  As String
    
    If (m_bFullScan) Then
        ReDim p_eClassRoots(0 To 1) As EScanRegistryRoots
        ReDim p_lScanDeeps(0 To 1) As Long
        ReDim p_sScanPaths(0 To 1) As String
        
        p_eClassRoots(0) = SCAN_LOCAL_MACHINE
        p_eClassRoots(1) = SCAN_USERS
        
        p_lScanDeeps(0) = 0
        p_lScanDeeps(1) = 0
        
        p_sScanPaths(0) = ""
        p_sScanPaths(1) = ""
    Else
        Select Case m_eCustomPath
            Case SCAN_ADDREMOVELISTS, SCAN_FILEEXTENSIONS, SCAN_HELPRESOURCES, SCAN_SHAREDDLLS, SCAN_SOFTWAREPATHS
                ReDim p_eClassRoots(0 To 0) As EScanRegistryRoots
                ReDim p_lScanDeeps(0 To 0) As Long
                ReDim p_sScanPaths(0 To 0) As String
                
                p_eClassRoots(0) = SCAN_LOCAL_MACHINE
                p_lScanDeeps(0) = 1
                
                Select Case m_eCustomPath
                    Case SCAN_ADDREMOVELISTS
                        p_sScanPaths(0) = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
                        
                    Case SCAN_FILEEXTENSIONS
                        p_eClassRoots(0) = SCAN_CLASSES_ROOT
                        p_sScanPaths(0) = ""
                        Me.FilterKeys = ".*|"
                        
                    Case SCAN_HELPRESOURCES
                        p_sScanPaths(0) = "Software\Microsoft\Windows\Help"
                        
                    Case SCAN_SHAREDDLLS
                        p_sScanPaths(0) = "Software\Microsoft\Windows\CurrentVersion\SharedDLLs"
                        
                    Case SCAN_SOFTWAREPATHS
                        p_sScanPaths(0) = "Software\Microsoft\Windows\CurrentVersion\App Paths"
                End Select
                
            Case SCAN_SHELLFOLDERS, SCAN_CUSTOMCONTROLS, SCAN_WINDOWSFONTS
                ReDim p_eClassRoots(0 To 1) As EScanRegistryRoots
                ReDim p_lScanDeeps(0 To 1) As Long
                ReDim p_sScanPaths(0 To 1) As String
                
                Select Case m_eCustomPath
                    Case SCAN_CUSTOMCONTROLS
                        p_eClassRoots(0) = SCAN_CLASSES_ROOT
                        p_eClassRoots(1) = SCAN_CLASSES_ROOT
                        
                        p_lScanDeeps(0) = 0
                        p_lScanDeeps(1) = 0
                        
                        p_sScanPaths(0) = "CLSID"
                        p_sScanPaths(1) = "TypeLib"
                        
                    Case SCAN_SHELLFOLDERS
                        p_eClassRoots(0) = SCAN_CURRENT_USER
                        p_eClassRoots(1) = SCAN_LOCAL_MACHINE
                        
                        p_lScanDeeps(0) = 1
                        p_lScanDeeps(1) = 1
                        
                        p_sScanPaths(0) = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
                        p_sScanPaths(1) = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
                        
                    Case SCAN_WINDOWSFONTS
                        p_eClassRoots(0) = SCAN_LOCAL_MACHINE
                        p_eClassRoots(1) = SCAN_LOCAL_MACHINE
                        
                        p_lScanDeeps(0) = 1
                        p_lScanDeeps(1) = 1
                        
                        p_sScanPaths(0) = "Software\Microsoft\Windows\CurrentVersion\Fonts"
                        p_sScanPaths(1) = "Software\Microsoft\Windows NT\CurrentVersion\Fonts"
                End Select
                
            Case SCAN_STARTUPKEYS
                ReDim p_eClassRoots(0 To 9) As EScanRegistryRoots
                ReDim p_lScanDeeps(0 To 9) As Long
                ReDim p_sScanPaths(0 To 9) As String
                
                For p_lCtr = 0 To 4
                    p_eClassRoots(p_lCtr) = SCAN_CURRENT_USER
                    p_eClassRoots(p_lCtr + 5) = SCAN_LOCAL_MACHINE
                    
                    p_lScanDeeps(p_lCtr) = 1
                    p_lScanDeeps(p_lCtr + 5) = 1
                    
                    Select Case p_lCtr
                        Case 0
                            p_sScanPaths(p_lCtr) = "Software\Microsoft\Windows\CurrentVersion\Run"
                            
                        Case 1
                            p_sScanPaths(p_lCtr) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
                            
                        Case 2
                            p_sScanPaths(p_lCtr) = "Software\Microsoft\Windows\CurrentVersion\RunServices"
                            
                        Case 3
                            p_sScanPaths(p_lCtr) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
                            
                        Case 4
                            p_sScanPaths(p_lCtr) = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
                    End Select
                    
                    p_sScanPaths(p_lCtr + 5) = p_sScanPaths(p_lCtr)
                Next
                
                p_sScanPaths(9) = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
        End Select
    End If
    
    StartCustomScanning p_eClassRoots, p_sScanPaths, p_lScanDeeps
End Sub

Private Sub StartCustomScanning( _
        ByRef p_ClassRoots() As EScanRegistryRoots, _
        ByRef p_ScanPaths() As String, _
        ByRef p_lScanDeeps() As Long)
        
    Dim p_lCtr      As Long
    Dim p_lUBound   As Long
    
    m_bCancel = False ' Ensure that we can still scan successfully
    m_bScanning = True
    p_lUBound = UBound(p_ClassRoots)
    
    For p_lCtr = 0 To p_lUBound
        If (m_bCancel) Then
            Exit Sub
        End If
        
        m_eClassRoot = p_ClassRoots(p_lCtr)
        m_lScanDeep = p_lScanDeeps(p_lCtr)
        m_sScanPath = p_ScanPaths(p_lCtr)
        
        InitializeScanning
        GetData m_sScanPath
        
        If (m_bScanSubKeys) Then
            GetKeys m_sScanPath
        End If
    Next
End Sub

Private Sub InitializeScanning()
    If (Not m_bCustomScan) Then
        m_bScanning = True
    End If
    
    m_bCancel = False ' Ensure that we can still scan successfully
    m_lStartDeep = KeyDeep(m_sScanPath)
End Sub

Private Sub FinalizeScanning()
    m_bScanning = False
    RaiseEvent DoneScanning(m_lTotalData, m_lTotalKeys)
    m_bCancel = False
End Sub

Private Function KeyDeep(Key As String) As Long
    Dim i As Integer
    
    If (Len(Key)) Then
        KeyDeep = 1
    End If
    
    Do
        i = InStr(i + 1, Key, "\")
        
        If (i > 0) Then
            KeyDeep = KeyDeep + 1 ' Count number of separators
        End If
    Loop While (i > 0)
End Function

' //-- Registry Scanning Procedures --//

Private Sub GetKeys(Path As String)
    If (m_bCancel) Then
        Exit Sub
    End If
    
    If (m_lScanDeep) Then
        If (KeyDeep(Path) - m_lStartDeep > m_lScanDeep) Then
            Exit Sub
        End If
    End If
    
    Dim p_hKey  As Long
    Dim p_lCtr  As Long
    Dim p_lRet  As Long
    Dim p_sPath As String
    
    If (m_bFilterKeys) Then
        If (KeyDeep(Path) - m_lStartDeep = 1) Then
            ' UCase$ for no case-sensitive filter search
            p_sPath = UCase$(Mid$(Path, InStrRev(Path, "\") + 1))
            
            For p_lCtr = 0 To m_lFilterKeys
                If (p_sPath Like m_sFilterKeys(p_lCtr)) Then
                    Exit For
                ElseIf (p_lCtr = m_lFilterKeys) Then
                    Exit Sub
                End If
            Next
        End If
    End If
    
    p_lRet = RegOpenKeyEx(m_eClassRoot, Path, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS, p_hKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        Dim p_lIndex    As Long
        Dim p_lMaxLen   As Long
        Dim p_sBuffer   As String
        Dim p_sKeys()   As String
        
        p_lCtr = 0
        p_lRet = RegQueryInfoKey(p_hKey, "", 0&, REG_OPTION_RESERVED, p_lCtr, p_lMaxLen, 0&, 0&, 0&, 0&, 0&, ByVal 0&)
        
        If (p_lRet = ERROR_SUCCESS) And (p_lCtr > 0) Then
            p_lCtr = p_lCtr - 1
            ReDim p_sKeys(0 To p_lCtr) As String
            p_lMaxLen = p_lMaxLen + 1 ' for the terminating null character
            
            Do
                If (GetQueueStatus(QS_ALLEVENTS)) Then
                    DoEvents ' Process system events other than this function
                End If
                
                If (m_bCancel) Then
                    RegCloseKey p_hKey
                    Exit Sub
                End If
                
                p_sBuffer = String$(p_lMaxLen, 0) ' Fill buffer with null characters
                p_lRet = RegEnumKey(p_hKey, p_lIndex, p_sBuffer, p_lMaxLen)
                
                If (p_lRet = ERROR_SUCCESS) Then
                    p_sBuffer = Left$(p_sBuffer, lstrlen(p_sBuffer))
                    
                    If (Len(Path) = 0) Then
                        p_sKeys(p_lIndex) = p_sBuffer
                    Else
                        p_sKeys(p_lIndex) = Path & "\" & p_sBuffer ' Store whole path
                    End If
                    
                    RaiseEvent BuildingKeyList(p_lIndex, p_lCtr)
                End If
                
                p_lIndex = p_lIndex + 1
            Loop While (p_lRet = ERROR_SUCCESS) And (p_lIndex < p_lCtr + 1)
        Else
            Exit Sub
        End If
        
        RegCloseKey p_hKey
    Else
        Exit Sub
    End If
    
    For p_lIndex = 0 To p_lCtr
        If (GetQueueStatus(QS_ALLEVENTS)) Then
            DoEvents ' Process system events other than this function
        End If
        
        If (m_bPaused) Then
            m_bScanning = False
            While (m_bPaused)
                DoEvents ' Don't execute next instructions until resumed
            Wend
            m_bScanning = True
        End If
        
        If (m_bCancel) Then
            Exit Sub
        End If
        
        p_sBuffer = p_sKeys(p_lIndex)
        
        If (Len(p_sBuffer)) Then  ' Avoid empty paths
            GetData p_sBuffer
            GetKeys p_sBuffer
        End If
    Next
End Sub

Private Sub GetData(Path As String)
    If (m_bCancel) Then
        Exit Sub
    End If
    
    If (m_lScanDeep) Then
        If (KeyDeep(Path) - m_lStartDeep > m_lScanDeep) Then
            Exit Sub
        End If
    End If
    
    Dim p_bDelete   As Boolean
    Dim p_hKey      As Long
    Dim p_lCtr      As Long
    Dim p_lRet      As Long
    Dim p_sRegPath  As String
    
    If (m_bFilterKeys) Then
        If (KeyDeep(Path) - m_lStartDeep = 1) Then
            ' UCase$ for no case-sensitive filter search
            p_sRegPath = UCase$(Mid$(Path, InStrRev(Path, "\") + 1))
            
            For p_lCtr = 0 To m_lFilterKeys
                If (p_sRegPath Like m_sFilterKeys(p_lCtr)) Then
                    Exit For
                ElseIf (p_lCtr = m_lFilterKeys) Then
                    Exit Sub
                End If
            Next
        End If
    End If
    
    p_bDelete = False
    m_lTotalKeys = m_lTotalKeys + 1
    RaiseEvent CurrentKey(Path, m_eClassRoot, p_bDelete)
    
    If (p_bDelete) Then
        m_bKeyDeleted = DeleteKey(m_eClassRoot, Path)
        
        If (m_bKeyDeleted) Then
            Exit Sub ' If successful, we don't need to continue below
        End If
    End If
    
    p_lRet = RegOpenKeyEx(m_eClassRoot, Path, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, p_hKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        Dim p_lIndex    As Long
        Dim p_lMaxLen   As Long
        Dim p_lSize     As Long
        Dim p_sBuffer   As String
        Dim p_sValues() As String
        
        p_lCtr = 0
        p_lRet = RegQueryInfoKey(p_hKey, "", 0&, REG_OPTION_RESERVED, 0&, 0&, 0&, p_lCtr, p_lMaxLen, 0&, 0&, ByVal 0&)
        
        If (p_lRet = ERROR_SUCCESS) And (p_lCtr > 0) Then
            p_lCtr = p_lCtr - 1
            ReDim p_sValues(0 To p_lCtr) As String
            
            Do
                If (GetQueueStatus(QS_ALLEVENTS)) Then
                    DoEvents ' Process system events other than this function
                End If
                
                If (m_bCancel) Then
                    RegCloseKey p_hKey
                    Exit Sub
                End If
                
                p_lSize = p_lMaxLen + 1
                p_sBuffer = String$(p_lSize, 0) ' Fill buffer with null characters
                p_lRet = RegEnumValue(p_hKey, (p_lIndex), p_sBuffer, p_lSize, REG_OPTION_RESERVED, 0&, 0&, 0&)
                
                If (p_lRet = ERROR_SUCCESS) Then
                    p_sValues(p_lIndex) = Left$(p_sBuffer, p_lSize)
                    
                    RaiseEvent BuildingDataList(p_lIndex, p_lCtr)
                End If
                
                p_lIndex = p_lIndex + 1
            Loop While (p_lRet = ERROR_SUCCESS) And (p_lIndex < p_lCtr + 1)
        Else
            Exit Sub
        End If
        
        RegCloseKey p_hKey
    Else
        Exit Sub
    End If
    
    For p_lIndex = 0 To p_lCtr
        If (GetQueueStatus(QS_ALLEVENTS)) Then
            DoEvents ' Process system events other than this function
        End If
        
        If (m_bPaused) Then
            m_bScanning = False
            While (m_bPaused)
                DoEvents ' Don't execute next instructions until resumed
            Wend
            m_bScanning = True
        End If
        
        If (m_bCancel) Then
            Exit Sub
        End If
        
        If (m_bFilterData) Then
            ' UCase$ for no case-sensitive filter search
            p_sRegPath = UCase$(p_sValues(p_lIndex))
            
            For p_lSize = 0 To m_lFilterData
                If (p_sRegPath Like m_sFilterData(p_lSize)) Then
                    Exit For
                Else
                    If (p_lSize = m_lFilterData) Then
                        GoTo Nxt_Data
                    End If
                End If
            Next
        End If
        
        p_sBuffer = p_sValues(p_lIndex)
        
        p_bDelete = False
        m_lTotalData = m_lTotalData + 1
        RaiseEvent CurrentData(p_sBuffer, Path, m_eClassRoot, p_bDelete)
        
        If (p_bDelete) Then
            m_bDataDeleted = DeleteValue(m_eClassRoot, Path, p_sBuffer)
        End If
Nxt_Data:
    Next
End Sub

' //-- Class Procedures --//

Private Sub Class_Initialize()
    ' Unless these values are set, these would be their default values
    m_eClassRoot = SCAN_CURRENT_USER
    m_lScanDeep = 0
    m_sScanPath = "Software\" & App.CompanyName & "\" & App.ProductName
End Sub

Private Sub Class_Terminate()
    If (m_bScanning Or m_bPaused) Then
        CancelScanning
    End If
End Sub

' Created by Noel A. Dacara | Copyright © 2003-2005 Davao City, Philippines
