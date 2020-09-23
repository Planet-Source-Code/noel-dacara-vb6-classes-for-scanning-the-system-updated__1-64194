VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScanDirectories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --------------------------------------------------------------------------------
' Name:     cScanDirectories class
' Author:   Noel A. Dacara (noeldacara@hotmail.com)
' Date:     02 August 2005
'
' You can freely use and distribute this file with or without permission provided
' that credits herein are kept unmodified and having acknowledgements is observed.
'
' Notes:    Only the 13 common special directories are in the custom scan paths.
'           But you can still specify other const CSIDL of special directories..
'           Note: Most CSIDL constant are for Windows® NT/2000, XP or later only
'
'           If ScanDrives is set, StartPath is ignored. If its optional argument
'           is specified, ScanDriveType is ignored. E.g ScanDrives("ACD") = True
'           Sample Filter Pattern: *.frm|*.bas|class1.cls|project1.vbp|module?.*
'
'           Filter *.* and * are not the same. *.* are files with file extension.
'
' Tips:     To attain the maximum scanning speed of your machine with this class,
'           you may need to close unnecessary programs running in your computer..
'
' Inquiries or comments regarding my works and whatsoever are greatly appreciated.
' --------------------------------------------------------------------------------

' API Declarations
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetQueueStatus Lib "user32.dll" (ByVal fuFlags As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function SHGetFolderPath Lib "ShFolder" Alias "SHGetFolderPathA" (ByVal hwnd As Long, ByVal ECP As Long, ByVal TOKENHANDLE As Long, ByVal flags As Long, ByVal lpPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

' API Constants
Private Const DRIVE_CDROM           As Long = 5
Private Const DRIVE_FIXED           As Long = 3
Private Const DRIVE_RAMDISK         As Long = 6
Private Const DRIVE_REMOTE          As Long = 4
Private Const DRIVE_REMOVABLE       As Long = 2
Private Const INVALID_HANDLE_VALUE  As Long = -1
Private Const MAX_PATH              As Long = 260
Private Const MAXDWORD              As Long = &HFFFFFFFF

' GetQueueStatus Flag
Private Const QS_ALLEVENTS As Long = &HBF

' API Types
Private Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Private Type SHORTITEMID
    cb      As Long
    abID    As Integer
End Type

Private Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

Private Type ITEMIDLIST
    mkid As SHORTITEMID
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

' User-Defined API Enum
Public Enum ESDCustomScanPaths
    ECP_DESKTOP = &H0                   ' Desktop
    ECP_PERSONAL = &H5                  ' My Documents
    ECP_STARTUP = &H7                   ' User startup
    ECP_RECENT = &H8                    ' Recent files
    ECP_STARTMENU = &HB                 ' Start menu
    ECP_FONTS = &H14                    ' Fonts
    ECP_COMMON_STARTMENU = &H16         ' All user start menu
    ECP_COMMON_STARTUP = &H18           ' All user startup
    ECP_COMMON_DESKTOPDIRECTORY = &H19  ' All user desktop
    ECP_WINDOWS = &H24                  ' Windows
    ECP_SYSTEM = &H25                   ' System
    ECP_PROGRAM_FILES = &H26            ' Program files
    ECP_TEMPORARYFILES = &HFF ' ???     ' User Temporary files
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private ECP_DESKTOP, ECP_PERSONAL, ECP_STARTUP
    Private ECP_RECENT, ECP_STARTMENU, ECP_FONTS, ECP_COMMON_STARTMENU
    Private ECP_COMMON_STARTUP, ECP_COMMON_DESKTOPDIRECTORY, ECP_WINDOWS
    Private ECP_SYSTEM, ECP_PROGRAM_FILES, ECP_TEMPORARYFILES
#End If

Public Enum ESDFileTimeType
    EFT_LASTWRITETIME = 0   ' Time the file was last modified
    EFT_CREATIONTIME = 1    ' Time the file was created
    EFT_LASTACCESSTIME = 2  ' Time the file was last accessed
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private EFT_CREATIONTIME, EFT_LASTACCESSTIME, EFT_LASTWRITETIME
#End If

Public Enum ESDDriveTypes
    EDT_FIXED = 2       ' Hard drives
    EDT_REMOVABLE = 4   ' Floppy drives, USB Flash drives...
    EDT_CDROM = 8       ' CD-ROM drives, DVD-ROMs...
    EDT_RAMDISK = 16    ' RAM disks
    EDT_REMOTE = 32     ' Network or remote drives
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private EDT_FIXED, EDT_REMOVABLE, EDT_CDROM, EDT_REMOTE, EDT_RAMDISK
#End If

' Variable Declarations
Private m_bCancel           As Boolean
Private m_bCustomScan       As Boolean
Private m_bErrorOnFail      As Boolean
Private m_bFileDeleted      As Boolean
Private m_bFiltered         As Boolean
Private m_bFolderDeleted    As Boolean
Private m_bPaused           As Boolean
Private m_bScanDrive        As Boolean
Private m_bScanning         As Boolean
Private m_bSubDirectories   As Boolean
Private m_eCustomPath       As ESDCustomScanPaths
Private m_eScanDriveType    As ESDDriveTypes
Private m_iScanDeep         As Integer
Private m_iStartDeep        As Integer
Private m_lAttributes       As VbFileAttribute
Private m_lFilterUBound     As Long
Private m_lTotalFiles       As Long
Private m_lTotalFolders     As Long
Private m_sCurrentDir       As String
Private m_sDriveLetters     As String
Private m_sFilters()        As String
Private m_sStartPath        As String
Private m_tFindData         As WIN32_FIND_DATA

' //-- Class Events --//

Public Event CurrentFile(File As String, Path As String, Delete As Boolean)
Attribute CurrentFile.VB_Description = "Occurs everytime a file is processed."
'Occurs everytime a file is processed.
Public Event CurrentFolder(Path As String, Cancel As Boolean, Delete As Boolean)
Attribute CurrentFolder.VB_Description = "Occurs when a folder is scanned for files."
'Occurs when a folder is scanned for files.
Public Event DoneScanning(TotalFolders As Long, TotalFiles As Long)
Attribute DoneScanning.VB_Description = "Occurs when scanning process is done."
'Occurs when scanning process is done.

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private File, Path, Delete, Cancel, TotalFolders, TotalFiles
#End If

' //-- Properties --//

Public Property Get Attributes() As VbFileAttribute
Attribute Attributes.VB_Description = "Returns/sets the attribute of files to be returned while scanning directories."
'Returns/sets the attribute of files to be returned while scanning directories.
    Attributes = m_lAttributes
End Property

Public Property Let Attributes(Value As VbFileAttribute)
    m_lAttributes = Value
End Property

Public Property Get CurrentFileScanned() As String
Attribute CurrentFileScanned.VB_Description = "Returns the filename of the current file scanned by the class."
'Returns the filename of the current file scanned by the class.
    CurrentFileScanned = m_tFindData.cFileName
End Property

Public Property Get CurrentFileAttribute() As VbFileAttribute
Attribute CurrentFileAttribute.VB_Description = "Returns the attribute of the current file/folder scanned by the class."
'Returns the attribute of the current file/folder scanned by the class.
    CurrentFileAttribute = m_tFindData.dwFileAttributes
End Property

Public Property Get CurrentFileSize() As Long
Attribute CurrentFileSize.VB_Description = "Returns the file size in bytes of the current file scanned by the class."
'Returns the file size in bytes of the current file scanned by the class.
    CurrentFileSize = (m_tFindData.nFileSizeHigh * MAXDWORD) + m_tFindData.nFileSizeLow
End Property

Public Property Get CurrentFileDate(Optional Mode As ESDFileTimeType) As Date
Attribute CurrentFileDate.VB_Description = "Returns the date property of the current file/folder scanned by the class."
'Returns the date property of the current file/folder scanned by the class.
    If (Len(m_tFindData.cFileName)) Then
        Dim p_tFileTime     As FILETIME
        Dim p_tSystemTime   As SYSTEMTIME
        
        Select Case Mode
            Case EFT_CREATIONTIME
                FileTimeToLocalFileTime m_tFindData.ftCreationTime, p_tFileTime
                
            Case EFT_LASTACCESSTIME
                FileTimeToLocalFileTime m_tFindData.ftLastAccessTime, p_tFileTime
                
            Case EFT_LASTWRITETIME
                FileTimeToLocalFileTime m_tFindData.ftLastWriteTime, p_tFileTime
        End Select
        
        FileTimeToSystemTime p_tFileTime, p_tSystemTime ' File time to system time
        
        CurrentFileDate = DateSerial(p_tSystemTime.wYear, _
                                     p_tSystemTime.wMonth, _
                                     p_tSystemTime.wDay)
    End If
End Property

Public Property Get CurrentFileTime(Optional Mode As ESDFileTimeType) As Date
Attribute CurrentFileTime.VB_Description = "Returns the time property of the current file/folder scanned by the class."
'Returns the time property of the current file/folder scanned by the class.
    If (Len(m_tFindData.cFileName)) Then
        Dim p_tFileTime     As FILETIME
        Dim p_tSystemTime   As SYSTEMTIME
        
        Select Case Mode
            Case EFT_CREATIONTIME
                FileTimeToLocalFileTime m_tFindData.ftCreationTime, p_tFileTime
                
            Case EFT_LASTACCESSTIME
                Exit Property ' Last accessed do not include time
                
            Case EFT_LASTWRITETIME
                FileTimeToLocalFileTime m_tFindData.ftLastWriteTime, p_tFileTime
        End Select
        
        FileTimeToSystemTime p_tFileTime, p_tSystemTime ' File time to system time
        
        CurrentFileTime = TimeSerial(p_tSystemTime.wHour, _
                                     p_tSystemTime.wMinute, _
                                     p_tSystemTime.wSecond)
    End If
End Property

Public Property Get CurrentFolderScanned() As String
Attribute CurrentFolderScanned.VB_Description = "Returns the current directory being processed."
'Returns the current directory being processed.
    CurrentFolderScanned = m_sCurrentDir
End Property

Public Property Get CustomScan() As Boolean
Attribute CustomScan.VB_Description = "Returns/sets whether to scan in the predefined special directories."
'Returns/sets whether to scan in the predefined special directories.
    CustomScan = m_bCustomScan
End Property

Public Property Let CustomScan(Value As Boolean)
    m_bCustomScan = Value
End Property

Public Property Get CustomScanPath() As ESDCustomScanPaths
Attribute CustomScanPath.VB_Description = "Returns/sets the custom scan path to scan for files and directory."
'Returns/sets the custom scan path to scan for files and directory.
    CustomScanPath = m_eCustomPath
End Property

Public Property Let CustomScanPath(Value As ESDCustomScanPaths)
    m_eCustomPath = Value
End Property

Public Property Get ErrorOnFail() As Boolean
Attribute ErrorOnFail.VB_Description = "Returns/sets a value whether to raise an error on failure to access path."
'Returns/sets a value whether to raise an error on failure to access path.
    ErrorOnFail = m_bErrorOnFail
End Property

Public Property Let ErrorOnFail(Value As Boolean)
    m_bErrorOnFail = Value
End Property

Public Property Get FileDeleted() As Boolean
Attribute FileDeleted.VB_Description = "Returns True if previous file with Delete argument is set on event is successfully removed."
'Returns True if previous file with Delete argument is set on event is successfully removed.
    FileDeleted = m_bFileDeleted
End Property

Public Property Get Filter() As String
Attribute Filter.VB_Description = "Returns/sets the filter patterns used to select files to be processed."
'Returns/sets the filter patterns used to select files to be processed.
    Filter = Join(m_sFilters(), "|") ' Return as one string
End Property

Public Property Let Filter(Value As String)
    If (Len(Value) = 0) Then
        Value = "*" ' Search for all files
    End If
    
    If (StrComp(Value, "*") = 0) Then
        m_bFiltered = False ' More faster!
    Else
        m_bFiltered = True
    End If
    
    m_sFilters = Split(UCase$(Value), "|") ' Save as an array of filter patterns
    m_lFilterUBound = UBound(m_sFilters)
End Property

Public Property Get FolderDeleted() As Boolean
Attribute FolderDeleted.VB_Description = "Returns True if previous folder with Delete argument is set on event is successfully removed."
'Returns True if previous folder with Delete argument is set on event is successfully removed.
    FolderDeleted = m_bFolderDeleted
End Property

Public Property Get Paused() As Boolean
Attribute Paused.VB_Description = "Returns True if scanning process has been temporarily stopped."
'Returns True if scanning process has been temporarily stopped.
    Paused = m_bPaused
End Property

Public Property Get ScanDeep() As Integer
Attribute ScanDeep.VB_Description = "Returns/sets a value to determine how deep from the startpath to scan."
'Returns/sets a value to determine how deep from the startpath to scan.
    ScanDeep = m_iScanDeep
End Property

Public Property Let ScanDeep(Value As Integer)
    m_iScanDeep = Value
End Property

Public Property Get ScanDrives(Optional DriveLetters As Variant) As Boolean
Attribute ScanDrives.VB_Description = "Returns/sets a value to determine if scanning is done through drives."
'Returns/sets a value to determine if scanning is done through drives.
    ScanDrives = m_bScanDrive
End Property

Public Property Let ScanDrives(Optional DriveLetters As Variant, Value As Boolean)
    If (IsMissing(DriveLetters)) Then
        m_sDriveLetters = ""
    Else
        m_sDriveLetters = DriveLetters
    End If
    
    m_bScanDrive = Value
End Property

Public Property Get ScanDriveType() As ESDDriveTypes
Attribute ScanDriveType.VB_Description = "Returns/sets a value to determine what type of drives to scan."
'Returns/sets a value to determine what type of drives to scan.
    ScanDriveType = m_eScanDriveType
End Property

Public Property Let ScanDriveType(Value As ESDDriveTypes)
    m_eScanDriveType = Value
End Property

Public Property Get Scanning() As Boolean
Attribute Scanning.VB_Description = "Returns True if class is currently on scanning state."
'Returns True if class is currently on scanning state.
    Scanning = m_bScanning
End Property

Public Property Get StartPath() As String
Attribute StartPath.VB_Description = "Returns/sets the path where to start scanning for files and folders."
'Returns/sets the path where to start scanning for files and folders.
    StartPath = m_sStartPath
End Property

Public Property Let StartPath(Value As String)
    m_sStartPath = Value
End Property

Public Property Get SubDirectories() As Boolean
Attribute SubDirectories.VB_Description = "Returns/sets whether to include subdirectories in scanning for files."
'Returns/sets whether to include subdirectories in scanning for files.
    SubDirectories = m_bSubDirectories
End Property

Public Property Let SubDirectories(Value As Boolean)
    m_bSubDirectories = Value
End Property

Public Property Get TotalFiles() As Long
Attribute TotalFiles.VB_Description = "Returns the current number of files processed."
'Returns the current number of files processed.
    TotalFiles = m_lTotalFiles
End Property

Public Property Get TotalFolders() As Long
Attribute TotalFolders.VB_Description = "Returns the current number of folders processed."
'Returns the current number of folders processed.
    TotalFolders = m_lTotalFolders
End Property

' //-- Procedures --//

Public Sub BeginScanning(Optional StartPath, Optional Filter)
Attribute BeginScanning.VB_Description = "Start scanning for the specified files according to filter pattern."
'Start scanning for the specified files according to filter pattern.
    If (m_bPaused) Then
        ResumeScanning ' Resume scanning instead
        Exit Sub
    End If
    
    If (m_bScanning) Then
        Exit Sub ' Avoid cascading scans
    End If
    
    m_lTotalFiles = 0 ' Reset number of files and folders processed
    m_lTotalFolders = 0
    
    If (m_bScanDrive) Then
        Dim i           As Long
        Dim p_lDrives   As Long
        Dim p_sDrives() As String
        Dim p_sBuffer   As String
        
        If (Len(m_sDriveLetters) = 0) Then
            p_sBuffer = Space$(MAX_PATH)
            i = GetLogicalDriveStrings(MAX_PATH, p_sBuffer)
            
            If (i) Then
                p_sBuffer = Left$(p_sBuffer, i)
                p_sBuffer = Replace$(p_sBuffer, "\", "")
                p_sDrives = Split(p_sBuffer, vbNullChar)
            Else
                Err.Raise 17 ' Can't perform requested operation
                Exit Sub
            End If
            
            p_lDrives = UBound(p_sDrives)
            
            For i = 0 To p_lDrives
                ' Filter drives
                If (GetDriveTypeEx(p_sDrives(i)) And m_eScanDriveType) Then
                    ' Just do nothing
                ElseIf (Len(p_sDrives(i))) Then
                    p_sBuffer = Replace$(p_sBuffer, p_sDrives(i) & vbNullChar, "")
                End If
            Next
            
            p_sDrives = Split(p_sBuffer, vbNullChar) ' Generate new filtered drive list
        Else
            p_lDrives = Len(m_sDriveLetters) - 1
            
            For i = 0 To p_lDrives
                ReDim Preserve p_sDrives(0 To i) As String
                p_sDrives(i) = Mid$(m_sDriveLetters, i + 1, 1) & ":"
            Next
        End If
        
        m_bCancel = False
        p_lDrives = UBound(p_sDrives)
        
        For i = 0 To p_lDrives
            If (Len(p_sDrives(i))) Then
                If (KeyExists(p_sDrives(i))) Then
                    StartScanning p_sDrives(i) ' Scan drive
                ElseIf (m_bErrorOnFail) Then
                    ' Usually occur on empty floppy or cd-rom drives
                    Err.Raise 76 ' Path not found
                End If
            End If
        Next i
        
        RaiseEvent DoneScanning(m_lTotalFolders, m_lTotalFiles)
    Else
        Dim p_StartPath As String
        
        If (Not IsMissing(Filter)) Then
            Me.Filter = Filter
        End If
        
        If (Not IsMissing(StartPath)) Then
            Me.StartPath = StartPath
        End If
        
        If (m_bCustomScan) Then
            p_StartPath = m_sStartPath ' Save startpath
            m_sStartPath = GetDirectory(Me.CustomScanPath)
        End If
        
        If (KeyExists(Me.StartPath)) Then ' Only scan if path exists
            m_bCancel = False
            StartScanning m_sStartPath
        ElseIf (m_bErrorOnFail) Then
            ' Usually occur on empty floppy or cd-rom drives
            Err.Raise 76 ' Path not found
        End If
        
        If (m_bCustomScan) Then
            m_sStartPath = p_StartPath ' Restore startpath
        End If
    End If
End Sub

Public Sub CancelScanning()
Attribute CancelScanning.VB_Description = "Abort the whole scanning process up to the current folder processed."
'Abort the whole scanning process up to the current folder processed.
    m_bCancel = True
    ResumeScanning ' Resume if scanning has been paused
End Sub

Public Sub PauseScanning()
Attribute PauseScanning.VB_Description = "Temporarily stop scanning process."
'Temporarily stop scanning process.
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

' //-- Private Procedures --//

Private Function GetDirectory(CustomScanPath As ESDCustomScanPaths) As String
    Dim p_Buffer    As String * MAX_PATH
    Dim t_IDList    As ITEMIDLIST
    Dim p_NullPos   As Integer
    Dim p_Result    As Long
    
    p_Result = SHGetSpecialFolderLocation(0&, CustomScanPath, t_IDList)
    
    If (p_Result) Then
        p_Result = SHGetFolderPath(0&, CustomScanPath, 0&, 0&, p_Buffer)
        
        If (p_Result = 0) Then
            p_Result = SHGetPathFromIDList(ByVal t_IDList.mkid.cb, ByVal p_Buffer)
        End If
    Else
        p_Result = SHGetPathFromIDList(ByVal t_IDList.mkid.cb, ByVal p_Buffer)
    End If
    
    GetDirectory = Left$(p_Buffer, lstrlen(p_Buffer))
    
    If (Len(GetDirectory) = 0) Then
        p_Buffer = String$(MAX_PATH, 0)
        
        Select Case CustomScanPath
            Case ECP_SYSTEM
                GetSystemDirectory p_Buffer, MAX_PATH
                
            Case ECP_TEMPORARYFILES
                GetTempPath MAX_PATH, p_Buffer
                
            Case ECP_WINDOWS, ECP_PROGRAM_FILES
                GetWindowsDirectory p_Buffer, MAX_PATH
                GetDirectory = Environ$("ProgramFiles") ' Not sure if everyone has this
        End Select
        
        If (CustomScanPath = ECP_PROGRAM_FILES) Then
            If (Len(GetDirectory) = 0) Then
                GetDirectory = Left$(p_Buffer, 3) & "Program Files"
            End If
        Else
            GetDirectory = Left$(p_Buffer, lstrlen(p_Buffer))
            
            If (GetDirectory Like "*\") Then
                GetDirectory = Left$(GetDirectory, Len(GetDirectory) - 1)
            End If
        End If
    End If
End Function

Private Function GetDriveTypeEx(ByVal nDrive As String) As ESDDriveTypes
    ' Our own GetDriveType function which supports AND & OR when used
    ' E.g. GetDriveTypeEx("C:") = (EDT_FIXED Or EDT_REMOVABLE)
    Select Case GetDriveType(nDrive)
        Case DRIVE_CDROM
            GetDriveTypeEx = EDT_CDROM
        Case DRIVE_FIXED
            GetDriveTypeEx = EDT_FIXED
        Case DRIVE_RAMDISK
            GetDriveTypeEx = EDT_RAMDISK
        Case DRIVE_REMOTE
            GetDriveTypeEx = EDT_REMOTE
        Case DRIVE_REMOVABLE
            GetDriveTypeEx = EDT_REMOVABLE
    End Select
End Function

Private Function GetShortPath(LongPath As String) As String
    Dim p_Buffer As String * MAX_PATH
    Dim p_Ret As Long
    
    p_Ret = GetShortPathName(LongPath, p_Buffer, MAX_PATH)
    GetShortPath = Left$(p_Buffer, p_Ret)
End Function

Private Function KeyDeep(Key As String) As Integer
    Dim i As Integer
    
    If (Len(Key)) Then
        KeyDeep = 1
    End If
    
    i = 1 ' Just an initialization
    
    While (i)
        i = InStr(i + 1, Key, "\")
        
        If (i) Then
            KeyDeep = KeyDeep + 1 'Count number of separators
        End If
    Wend
End Function

Private Function KeyExists(ByVal Path As String) As Boolean
    If (Not Path Like "*\") Then
        Path = Path & "\"
    End If
    
    Path = Path & "nul"
    
    If (Len(Dir$(Path))) Then
        KeyExists = True ' Path/folder exists...
    End If
End Function

Private Sub StartScanning(Path As String)
    If (m_lAttributes And vbDirectory) = vbDirectory Then
        m_lAttributes = m_lAttributes - vbDirectory ' Must be a file
    End If
    
    Dim p_lFind As Long
    
    m_iStartDeep = KeyDeep(Path)
    m_bScanning = True
    
    ' Get folder information about the start path
    If (Len(Path) - 2 > MAX_PATH) Then
        p_lFind = FindFirstFile(GetShortPath(Path & "\*"), m_tFindData)
    Else
        p_lFind = FindFirstFile(Path & "\*", m_tFindData)
    End If
    
    If (p_lFind <> INVALID_HANDLE_VALUE) Then
        FindClose p_lFind ' Close search handle immediately
    End If
    
    ScanForFiles Path ' Start scanning
    
    m_bScanning = False
    
    If (Not m_bScanDrive) Then
        RaiseEvent DoneScanning(m_lTotalFolders, m_lTotalFiles)
    End If
End Sub

Private Sub ScanForFiles(Path As String)
    If (m_bCancel) Then
        Exit Sub 'Abort scanning
    End If
    
    On Error Resume Next
    
    Dim i           As Long
    Dim j           As Long
    Dim p_bDelete   As Boolean
    Dim p_lFind     As Long
    Dim p_lRet      As Long
    Dim p_sData     As String
    
    If (m_iScanDeep) Then
        If (KeyDeep(Path) - m_iStartDeep > m_iScanDeep) Then
            Exit Sub
        End If
    End If
    
    m_sCurrentDir = Path
    m_lTotalFolders = m_lTotalFolders + 1
    RaiseEvent CurrentFolder(Path, m_bCancel, p_bDelete) ' Send folder to event
    Path = m_sCurrentDir
    
    If (p_bDelete) Then
        ' Attempt to remove directory
        m_bFolderDeleted = (RemoveDirectory(Path) <> 0)
        p_bDelete = False
        
        If (m_bFolderDeleted) Then
            Exit Sub
        End If
    End If
    
    If (Len(Path) - 2 > MAX_PATH) Then
        p_lFind = FindFirstFile(GetShortPath(Path & "\*"), m_tFindData)
    Else
        p_lFind = FindFirstFile(Path & "\*", m_tFindData)
    End If
    
    If (p_lFind <> INVALID_HANDLE_VALUE) Then
        p_lRet = p_lFind
        
        While (p_lRet)
            If (GetQueueStatus(QS_ALLEVENTS)) Then  ' Check for events
                DoEvents ' Process system events
            End If
            
            If (m_bPaused) Then
                m_bScanning = False
                While (m_bPaused)
                    DoEvents ' Don't execute next instruction until resumed
                Wend
                m_bScanning = True
            End If
            
            If (m_bCancel) Then
                FindClose p_lFind ' Close search handle
                Exit Sub ' Abort scanning
            End If
            
            p_sData = m_tFindData.cFileName
            p_sData = Left$(p_sData, lstrlen(p_sData))
            
            If (j <> 3) Then
                ' Ignore the first two items
                ' .
                ' ..
                j = j + 1
            End If
            
            If (j = 3) Then
                If (m_tFindData.dwFileAttributes And vbDirectory) = vbDirectory Then
                    If (m_bSubDirectories) Then
                        ' Process sub directory immediately
                        ScanForFiles Path & "\" & p_sData
                    End If
                ElseIf (m_tFindData.dwFileAttributes Or m_lAttributes) Then
                    If (Not m_bFiltered) Then
                        m_lTotalFiles = m_lTotalFiles + 1
                        ' Send file to event
                        RaiseEvent CurrentFile(p_sData, Path, p_bDelete)
                    Else
                        For i = 0 To m_lFilterUBound
                            If (UCase$(p_sData) Like m_sFilters(i)) Then
                                m_lTotalFiles = m_lTotalFiles + 1
                                ' Send file to event
                                RaiseEvent CurrentFile(p_sData, Path, p_bDelete)
                            End If
                        Next
                    End If
                    
                    If (p_bDelete) Then
                        p_bDelete = False
                        ' Attemp to delete file
                        m_bFileDeleted = (DeleteFile(Path & "\" & p_sData) <> 0)
                    End If
                End If
            End If
            
            p_lRet = FindNextFile(p_lFind, m_tFindData)
        Wend
        
        FindClose p_lFind ' Close search handle
    End If
End Sub

' //-- Class Procedures --//

Private Sub Class_Initialize()
    ' Unless these values are set, these would be their default values
    m_eScanDriveType = EDT_FIXED Or EDT_REMOVABLE
    m_lAttributes = vbNormal Or vbHidden Or vbReadOnly Or vbArchive Or vbSystem
    Filter = "*"
    m_sStartPath = CurDir$()
    m_bSubDirectories = True
End Sub

Private Sub Class_Terminate()
    If (m_bScanning Or m_bPaused) Then
        CancelScanning ' Just to ensure that everything will be alright
    End If
End Sub

' Created by Noel A. Dacara | Copyright © 2003-2005 Davao City, Philippines
