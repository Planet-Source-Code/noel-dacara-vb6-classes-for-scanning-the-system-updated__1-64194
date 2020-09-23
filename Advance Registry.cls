VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAdvanceRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --------------------------------------------------------------------------------
' Name:     cAdvanceRegistry class
' Author:   Noel A. Dacara (noeldacara@hotmail.com)
' Date:     2 June 2005
'
' You can freely use and distribute this file with or without permission provided
' that credits herein are kept unmodified and having acknowledgements is observed.
'
' Caution:  Incorrectly editing the registry may severely damage your system.
'           Before using this class, backup any valued data on your computer.
'
' Notes:    To access Remote registry for Windows NT 4.0, 2000, XP, or 2003, run
'           the program from an account that has permission to that remote system.
'
'           Be careful in using DeleteTree function. Backup registry for assurance.
'           Author will not be held liable for any damages while using the class.
'
'           Binary typed registry data are returned as an array of ascii numbers.
'           Binary in unicode format can be usually read as normal strings.
'
' Sample:   Dim KeyArray() As String
'           If REG.EnumerateKeys(KeyArray, HKEY_LOCAL_MACHINE, "Software") Then
'               For i = 0 To UBound(KeyArray)
'                   Debug.Print KeyArray(i)
'               Next
'           End If
'
' Inquiries or comments regarding my works and whatsoever are greatly appreciated.
' --------------------------------------------------------------------------------

' API declarations
Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetQueueStatus Lib "user32.dll" (ByVal fuFlags As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, ByRef phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Modified API declarations
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Long, ByRef Source As Long, ByVal Length As Long)
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Byte, ByVal lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByVal lpReserved As Long, ByRef lpcSubKeys As Long, ByRef lpcbMaxSubKeyLen As Long, ByRef lpcbMaxClassLen As Long, ByRef lpcValues As Long, ByRef lpcbMaxValueNameLen As Long, ByRef lpcbMaxValueLen As Long, ByRef lpcbSecurityDescriptor As Long, ByRef lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lKey As Long, ByVal lpValueName As String, ByVal lReserved As Long, ByRef lpType As Long, ByRef lpData As Byte, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExDouble Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lKey As Long, ByVal lpValueName As String, ByVal lReserved As Long, ByRef lpType As Long, ByRef lpData As Double, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lKey As Long, ByVal lpValueName As String, ByVal lReserved As Long, ByRef lpType As Long, ByRef lpData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lKey As Long, ByVal lpValueName As String, ByVal lReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExDouble Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Double, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

' API Constants
Private Const QS_HOTKEY                 As Long = &H80
Private Const QS_KEY                    As Long = &H1
Private Const QS_MOUSEBUTTON            As Long = &H4
Private Const QS_MOUSEMOVE              As Long = &H2
Private Const QS_MOUSE                  As Long = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT                  As Long = (QS_MOUSE Or QS_KEY)
Private Const QS_PAINT                  As Long = &H20
Private Const QS_POSTMESSAGE            As Long = &H8
Private Const QS_TIMER                  As Long = &H10
Private Const QS_ALLEVENTS              As Long = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)

' Registry security option constants
Private Const KEY_CREATE_LINK           As Long = &H20
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_SET_VALUE             As Long = &H2
Private Const STANDARD_RIGHTS_ALL       As Long = &H1F0000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const KEY_ALL_ACCESS            As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

' Registry create type values
Private Const REG_OPTION_NON_VOLATILE   As Long = 0 ' Key is preserved on system reboot
Private Const REG_OPTION_RESERVED       As Long = 0 ' Parameter is reserved for future use

' Registry create/open disposition
Private Const REG_CREATED_NEW_KEY   As Long = &H1   ' New registry key was created
Private Const MAX_CLASS_NAME        As Long = 255

' Registry return constants
Private Const ERROR_ACCESS_DENIED   As Long = 5&    ' Access to registry key was denied
Private Const ERROR_MORE_DATA       As Long = 234   ' More data is available
Private Const ERROR_NO_MORE_ITEMS   As Long = 259&  ' No data is available
Private Const ERROR_SUCCESS         As Long = 0&    ' Operation was completed successfully

' API types
Private Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength                 As Long
    lpSecurityDescriptor    As Long
    bInheritHandle          As Long
End Type

' Public enums
Public Enum ERegistryRoots
    HKEY_CLASSES_ROOT = &H80000000  ' Contains file association mappings
    HKEY_CURRENT_USER = &H80000001  ' Contains information about the current user
    HKEY_LOCAL_MACHINE = &H80000002 ' Contains computer specific information
    HKEY_USERS = &H80000003         ' Contains individual preferences for each users
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS
#End If

'Common registry data types
Public Enum ERegistryDataTypes
    REG_BINARY = 3                  ' Binary data in any form
    REG_DWORD = 4                   ' 32-bit number
    REG_DWORD_BIG_ENDIAN = 5        ' 32-bit number in big-endian format
    REG_DWORD_LITTLE_ENDIAN = 4     ' 32-bit number in little-endian format
    REG_EXPAND_SZ = 2               ' Contains unexpanded references to environment variables (Ex. %SystemRoot%=C:\Windows)
    REG_MULTI_SZ = 7                ' String that contain lists or multiple values
    REG_QWORD = 11                  ' 64-bit number
    REG_QWORD_LITTLE_ENDIAN = 11    ' 64-bit number in little-endian format
    REG_SZ = 1                      ' Standard string
    
    ' Read-only purpose data types
    REG_FULL_RESOURCE_DESCRIPTOR = 9    '
    REG_LINK = 6                        ' Unicode symbolic link
    REG_NONE = 0                        ' No defined value type
    REG_RESOURCE_LIST = 8               ' A device-driver resource list
    REG_RESOURCE_REQUIREMENTS_LIST = 10 '
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private REG_BINARY, REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
    Private REG_EXPAND_SZ, REG_MULTI_SZ, REG_QWORD, REG_QWORD_LITTLE_ENDIAN
    Private REG_SZ, REG_FULL_RESOURCE_DESCRIPTOR, REG_LINK, REG_NONE
    Private REG_RESOURCE_LIST, REG_RESOURCE_REQUIREMENTS_LIST
#End If

' Variable declarations
Private m_bRemoteConnection As Boolean
Private m_eClassRoot        As ERegistryRoots
Private m_eDataType         As ERegistryDataTypes
Private m_eRemoteClassRoot  As ERegistryRoots
Private m_lRemoteKey        As Long
Private m_sKey              As String
Private m_sMachineName      As String
Private m_sValueName        As String
Private m_vDefault          As Variant

' //-- Properties --//

Public Property Get ClassRoot() As ERegistryRoots
Attribute ClassRoot.VB_Description = "Returns/sets a value to determine the registry key root type used."
'Returns/sets a value to determine the registry key root type used.
    ClassRoot = m_eClassRoot
End Property

Public Property Let ClassRoot(Value As ERegistryRoots)
    If (Value = 0) Then
        Err.Raise -1, , "Invalid registry class root value."
        Exit Property
    End If
    
    m_eClassRoot = Value
End Property

Public Property Get DataType() As ERegistryDataTypes
Attribute DataType.VB_Description = "Returns a value to determine the data type used or set the data type to be used."
'Returns a value to determine the data type used or set the data type to be used.
    DataType = m_eDataType
End Property

Public Property Let DataType(Value As ERegistryDataTypes)
    m_eDataType = Value
End Property

Public Property Get Default() As Variant
Attribute Default.VB_Description = "Returns/sets the default registry value to be used incase of problems."
'Returns/sets the default registry value to be used incase of problems.
    Default = m_vDefault
End Property

Public Property Let Default(Value As Variant)
    m_vDefault = Value
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets a value to determines the registry key/section to be used."
'Returns/sets a value to determines the registry key/section to be used.
    Key = m_sKey
End Property

Public Property Let Key(Value As String)
    m_sKey = Value
End Property

Public Property Get MachineName() As String
Attribute MachineName.VB_Description = "Returns/sets the name of computer where to establish a remote registry connection."
'Returns/sets the name of computer where to establish a remote registry connection.
    MachineName = m_sMachineName
End Property

Public Property Let MachineName(Value As String)
    ValidateMachineName Value ' make sure name is in the proper format
    m_sMachineName = Value
End Property

Public Property Get Value() As Variant
Attribute Value.VB_Description = "Returns/sets the value of a registry data."
'Returns/sets the value of a registry data.
    Value = Me.ValueEx
End Property

Public Property Let Value(Value As Variant)
    Me.ValueEx = Value
End Property

Public Property Get ValueEx( _
        Optional ClassRoot, _
        Optional Key, _
        Optional ValueName) As Variant
Attribute ValueEx.VB_Description = "Returns/sets the value of the specified registry data."
'Returns/sets the value of the specified registry data.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    If (Not IsMissing(ValueName)) Then
        m_sValueName = ValueName
    End If
    
    Dim p_bBinary() As Byte
    Dim p_lDword    As Long
    Dim p_lKey      As Long
    Dim p_lRet      As Long
    Dim p_lDataLen  As Long
    Dim p_lDataType As Long
    Dim p_sBuffer   As String
    
    If (m_bRemoteConnection) Then
        ClassRoot = m_lRemoteKey
    End If
    
    p_lRet = RegOpenKeyEx(m_eClassRoot, m_sKey, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, p_lKey)
    p_lRet = RegQueryValueExLong(p_lKey, m_sValueName, REG_OPTION_RESERVED, p_lDataType, 0&, p_lDataLen)
    
    If (p_lRet) And (p_lRet <> ERROR_MORE_DATA) Then
        ValueEx = m_vDefault
        Exit Property
    End If
    
    m_eDataType = p_lDataType
    
    Select Case p_lDataType
        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
            p_sBuffer = String$(p_lDataLen, 0)
            p_lRet = RegQueryValueExString(p_lKey, m_sValueName, REG_OPTION_RESERVED, p_lDataType, p_sBuffer, p_lDataLen)
            
            If (p_lDataType = REG_EXPAND_SZ) Then
                ValueEx = ExpandEnvironmentString(p_sBuffer) ' Expand %VARIABLES%
            Else
                If (p_lDataLen > 0) Then
                    ValueEx = Left$(p_sBuffer, p_lDataLen - 1)
                End If
            End If
            
        Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN, REG_DWORD_BIG_ENDIAN
            p_lRet = RegQueryValueExLong(p_lKey, m_sValueName, REG_OPTION_RESERVED, p_lDataType, p_lDword, p_lDataLen)
            
            If (p_lDataType = REG_DWORD) Then
                ValueEx = CLng(p_lDword)
            Else
                ValueEx = GetBigEndianValue(p_lDword) ' REG_DWORD_BIG_ENDIAN
            End If
            
        Case REG_QWORD, REG_QWORD_LITTLE_ENDIAN
            Dim p_dQword As Double
            
            p_lRet = RegQueryValueExDouble(p_lKey, m_sValueName, REG_OPTION_RESERVED, p_lDataType, p_dQword, p_lDataLen)
            ValueEx = CDbl(p_dQword)
            
        Case Else
            If (p_lDataLen > 0) Then
                ReDim p_bBinary(p_lDataLen - 1) ' without the terminating null character
                
                p_lRet = RegQueryValueExByte(p_lKey, m_sValueName, REG_OPTION_RESERVED, p_lDataType, p_bBinary(0), p_lDataLen)
            End If
            
            ValueEx = p_bBinary ' return as an array of characters
    End Select
    
    RegCloseKey p_lKey
End Property

Public Property Let ValueEx( _
        Optional ClassRoot, _
        Optional Key, _
        Optional ValueName, _
        Value As Variant)
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    If (Not IsMissing(ValueName)) Then
        m_sValueName = ValueName
    End If
    
    Dim p_bByte()       As Byte
    Dim p_dQword        As Double
    Dim p_lDword        As Long
    Dim p_lKey          As Long
    Dim p_lLen          As Long
    Dim p_lRet          As Long
    Dim p_sString       As String
    Dim p_tSecAttrib    As SECURITY_ATTRIBUTES
    
    If (m_bRemoteConnection) Then
        m_eClassRoot = m_lRemoteKey
    End If
    
    p_lRet = RegCreateKeyEx(m_eClassRoot, m_sKey, REG_OPTION_RESERVED, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_tSecAttrib, p_lKey, 0&)
    
    If (p_lRet = ERROR_SUCCESS) Then
        Select Case m_eDataType
            Case REG_BINARY
                ' Can also accept strings
                If (VarType(Value) = (vbArray + vbByte) Or vbString) Then
                    p_bByte = Value
                Else
                    Err.Raise m_eDataType, , "Cannot set the specified value in the registry using the defined data type."
                End If
                
                p_lLen = UBound(p_bByte) - LBound(p_bByte) + 1
                p_lRet = RegSetValueExByte(p_lKey, m_sValueName, REG_OPTION_RESERVED, m_eDataType, p_bByte(0), p_lLen)
                
            Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
                p_sString = Value
                
                If (m_eDataType = REG_MULTI_SZ) Then
                    ' Replace all linefeeds with null characters
                    p_sString = Replace$(p_sString, vbNewLine, vbNullChar)
                    
                    ' Windows® Registry Editor does not allow REG_MULTI_SZ
                    ' typed data to contain empty lines and so as we...
                    While InStr(1, p_sString, String$(2, vbNullChar)) > 0
                        p_sString = Replace$(p_sString, String$(2, vbNullChar), vbNullChar)
                    Wend
                    
                    ' REG_MULTI_SZ typed data is terminated by two null characters
                    If Right$(p_sString, 1) = vbNullChar Then
                        p_sString = p_sString & vbNullChar
                    End If
                End If
                
                p_sString = p_sString & vbNullChar ' Terminate with null character
                p_lLen = Len(p_sString)
                
                p_lRet = RegSetValueExString(p_lKey, m_sValueName, REG_OPTION_RESERVED, m_eDataType, p_sString, p_lLen)
                
            Case REG_DWORD, 5 ' REG_DWORD_LITTLE_ENDIAN, REG_DWORD_BIG_ENDIAN
                If (VarType(Value) = vbInteger Or vbLong) Then
                    p_lDword = CLng(Value)
                    p_lLen = 4 ' 32-bits
                    
                    p_lRet = RegSetValueExLong(p_lKey, m_sValueName, REG_OPTION_RESERVED, m_eDataType, p_lDword, p_lLen)
                Else
                    Err.Raise m_eDataType, , "Cannot set the specified value in the registry using the defined data type."
                End If
                
            Case REG_QWORD ' REG_QWORD_LITTLE_ENDIAN
                If (VarType(Value) = vbInteger Or vbLong Or vbDouble) Then
                    p_dQword = CDbl(Value)
                    p_lLen = 8 ' 64-bits
                    
                    p_lRet = RegSetValueExDouble(p_lKey, m_sValueName, REG_OPTION_RESERVED, m_eDataType, p_dQword, p_lLen)
                Else
                    Err.Raise m_eDataType, , "Cannot set the specified value in the registry using the defined data type."
                End If
                
            Case Else
                Err.Raise m_eDataType, , "The specified data type is either invalid or not supported for write purpose."
        End Select
        
        RegCloseKey p_lKey
    Else
        Err.Raise ClassRoot, , "Unable to open/create registry key: '" & m_sKey & "' for setting new value of '" & m_sValueName & "' to '" & Value & "'"
    End If
End Property

Public Property Get ValueName() As String
Attribute ValueName.VB_Description = "Returns/sets a name to identify a particular registry data."
'Returns/sets a name to identify a particular registry data.
    ValueName = m_sValueName
End Property

Public Property Let ValueName(Value As String)
    m_sValueName = Value
End Property

' //-- Public procedures --//

Public Function CreateKey(Optional ClassRoot, Optional Key) As Boolean
Attribute CreateKey.VB_Description = "Creates a new registry key from a specified path of a registry root."
'Creates a new registry key from a specified path of a registry root.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    Dim p_lKey          As Long
    Dim p_lRet          As Long
    Dim p_tSecAttrib    As SECURITY_ATTRIBUTES
    
    If (m_bRemoteConnection) Then
        m_eClassRoot = m_lRemoteKey
    End If
    
    p_lRet = RegCreateKeyEx(m_eClassRoot, m_sKey, REG_OPTION_RESERVED, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_tSecAttrib, p_lKey, 0&)
    
    If (p_lRet = ERROR_SUCCESS) Then
        CreateKey = True
        RegCloseKey p_lKey
    Else
        Err.Raise ClassRoot, , "Unable to create registry key: '" & m_sKey & "'"
    End If
End Function

Public Function DeleteAllData(Optional ClassRoot, Optional Key) As Boolean
Attribute DeleteAllData.VB_Description = "Remove and clear all data of the specified registry key."
'Remove and clear all data of the specified registry key.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    Dim p_sDataArray()  As String
    Dim p_lCtr          As Long
    
    If (EnumerateData(p_sDataArray, m_eClassRoot, m_sKey)) Then
        For p_lCtr = LBound(p_sDataArray) To UBound(p_sDataArray)
            DeleteData m_eClassRoot, m_sKey, p_sDataArray(p_lCtr)
        Next
        
        DeleteAllData = True
    End If
End Function

Public Function DeleteData(Optional ClassRoot, Optional Key, Optional ValueName) As Boolean
Attribute DeleteData.VB_Description = "Removes the specified data on a particular registry key and class root."
'Removes the specified data on a particular registry key and class root.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    If (Not IsMissing(ValueName)) Then
        m_sValueName = ValueName
    End If
    
    Dim p_lKey As Long
    Dim p_lRet As Long
    
    If (m_bRemoteConnection) Then
        m_eClassRoot = m_lRemoteKey
    End If
    
    p_lRet = RegOpenKeyEx(m_eClassRoot, m_sKey, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_lKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        p_lRet = RegDeleteValue(p_lKey, m_sValueName)
        
        If (p_lRet = ERROR_SUCCESS) Then
            DeleteData = True
        End If
        
        RegCloseKey p_lKey
    Else
        Err.Raise ClassRoot, , "Unable to open registry key '" & m_sKey & "' for delete access."
    End If
End Function

Public Function DeleteKey(Optional ClassRoot, Optional Key) As Boolean
Attribute DeleteKey.VB_Description = "Removes the specified key on a particular class root in the registry."
'Removes the specified key on a particular class root in the registry.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    Dim p_lKey          As Long
    Dim p_lRet          As Long
    Dim p_sChildKey     As String
    Dim p_sParentKey    As String
    
    If (m_bRemoteConnection) Then
        ClassRoot = m_lRemoteKey
    End If
    
    p_sParentKey = GetParentKey(m_sKey) ' Get parent key
    p_lRet = RegOpenKeyEx(m_eClassRoot, p_sParentKey, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_lKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        If (Len(p_sParentKey) = 0) Then
            p_sParentKey = m_sKey
        Else
            p_sChildKey = Mid$(m_sKey, Len(p_sParentKey) + 2)
        End If
        
        p_lRet = RegDeleteKey(p_lKey, p_sChildKey)
        
        If (p_lRet = ERROR_SUCCESS) Then
            DeleteKey = True
        End If
        
        RegCloseKey p_lKey
    Else
        Err.Raise ClassRoot, , "Unable to open registry key '" & p_sParentKey & "' for delete access."
    End If
End Function

Public Function DeleteTree(Optional ClassRoot, Optional Key) As Boolean
Attribute DeleteTree.VB_Description = "Remove the specified registry key and all of its subkeys."
'Remove the specified registry key and all of its subkeys.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    Dim p_bHasSubKey    As Boolean
    Dim p_sKeyArray()   As String
    Dim p_sCurrentKey   As String
    
    p_sCurrentKey = m_sKey ' Set current key
    
    Do
        If (p_bHasSubKey) Then
            p_sCurrentKey = p_sCurrentKey & "\" & p_sKeyArray(0) 'Go deeper
        Else
            ' Check if current key is a subkey
            If (StrComp(p_sCurrentKey, m_sKey) <> 0) Then
                p_sCurrentKey = GetParentKey(p_sCurrentKey)
            End If
        End If
        
        ' Ensure that the system still process other events
        If (GetQueueStatus(QS_ALLEVENTS)) Then
            DoEvents
        End If
        
        DeleteSubKeys m_eClassRoot, p_sCurrentKey
        
        ' Determine and get if current key has any subkeys
        p_bHasSubKey = Me.EnumerateKeys(p_sKeyArray, m_eClassRoot, p_sCurrentKey)
        
        If (p_sCurrentKey = m_sKey) And (Not p_bHasSubKey) Then
            Exit Do ' This is where we get out the loop
        End If
    Loop ' Does it looks like an infinite loop? hehe...
    
    If (DeleteKey(m_eClassRoot, m_sKey)) Then ' Attempt to delete the main key
        DeleteTree = True ' If successful, then so as this function...
    End If
End Function

Public Function EnumerateData( _
        DataArray() As String, _
        Optional ClassRoot, _
        Optional Key) As Boolean
'Enumerate the list of valuenames found on a specified registry key.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    Dim p_lKey As Long
    Dim p_lRet As Long
    
    If (m_bRemoteConnection) Then
        m_eClassRoot = m_lRemoteKey
    End If
    
    p_lRet = RegOpenKeyEx(m_eClassRoot, m_sKey, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, p_lKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        Dim p_sBuffer       As String
        Dim p_lDataCount    As Long
        Dim p_lDataLen      As Long
        Dim p_lIndex        As Long
        Dim p_lMaxLen       As Long
        
        Dim FT As FILETIME
        p_lRet = RegQueryInfoKey(p_lKey, "", 0&, REG_OPTION_RESERVED, 0&, 0&, 0&, p_lDataCount, p_lMaxLen, 0&, 0&, FT)
        
        If (p_lDataCount > 0) Then
            ReDim DataArray(0 To (p_lDataCount - 1)) As String
            
            Do
                p_lDataLen = p_lMaxLen + 1
                p_sBuffer = String$(p_lDataLen, 0)
                p_lRet = RegEnumValue(p_lKey, p_lIndex, p_sBuffer, p_lDataLen, REG_OPTION_RESERVED, 0&, 0&, 0&)
                
                If (p_lRet = ERROR_SUCCESS) Then
                    DataArray(p_lIndex) = Left$(p_sBuffer, p_lDataLen)
                End If
                
                p_lIndex = p_lIndex + 1
            Loop While ((p_lRet = ERROR_SUCCESS)) And (p_lIndex < p_lDataCount)
            
            EnumerateData = True
        End If
        
        RegCloseKey p_lKey
    Else
        Err.Raise ClassRoot, , "Unable to open registry key '" & m_sKey & "' for data query access."
    End If
End Function

Public Function EnumerateKeys( _
        KeyArray() As String, _
        Optional ClassRoot, _
        Optional Key) As Boolean
'Enumerate the list of subkeys found on a specified registry key.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    Dim p_lKey As Long
    Dim p_lRet As Long
    
    If (m_bRemoteConnection) Then
        m_eClassRoot = m_lRemoteKey
    End If
    
    p_lRet = RegOpenKeyEx(m_eClassRoot, m_sKey, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS, p_lKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        Dim p_lIndex    As Long
        Dim p_lKeyCount As Long
        Dim p_lMaxLen   As Long
        Dim p_sBuffer   As String
        
        Dim FT As FILETIME
        p_lRet = RegQueryInfoKey(p_lKey, "", 0&, REG_OPTION_RESERVED, p_lKeyCount, p_lMaxLen, 0&, 0&, 0&, 0&, 0&, FT)
        
        If (p_lKeyCount > 0) Then
            ReDim KeyArray(0 To (p_lKeyCount - 1)) As String
            p_lMaxLen = p_lMaxLen + 1 ' for the terminating null character
            
            Do
                p_sBuffer = String$(p_lMaxLen, 0)
                p_lRet = RegEnumKey(p_lKey, p_lIndex, p_sBuffer, p_lMaxLen)
                
                If (p_lRet = ERROR_SUCCESS) Then
                    KeyArray(p_lIndex) = Left$(p_sBuffer, lstrlen(p_sBuffer))
                End If
                
                p_lIndex = p_lIndex + 1
            Loop While ((p_lRet = ERROR_SUCCESS)) And (p_lIndex < p_lKeyCount)
            
            EnumerateKeys = True
        End If
        
        RegCloseKey p_lKey
    Else
        Err.Raise ClassRoot, , "Unable to open registry key '" & Key & "' for key query access."
    End If
End Function

Public Sub Export( _
        File As String, _
        Optional Overwrite As Boolean, _
        Optional ClassRoot, _
        Optional Key)
'Attempt to export the contents of a registry key and all of its subkeys to a file.
    If (Not Overwrite) Then
        If (Len(Dir$(File))) Then
            Exit Sub ' Do not overwrite existing file
        End If
    End If
    
    Dim p_sParam As String
    
    If (Not IsMissing(ClassRoot)) Then
        p_sParam = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    If (IsNumeric(p_sParam)) Then
        Select Case ClassRoot
            Case HKEY_CLASSES_ROOT
                p_sParam = "HKEY_CLASSES_ROOT"
            Case HKEY_CURRENT_USER
                p_sParam = "HKEY_CURRENT_USER"
            Case HKEY_LOCAL_MACHINE
                p_sParam = "HKEY_LOCAL_MACHINE"
            Case HKEY_USERS
                p_sParam = "HKEY_USERS"
            Case 0
                Err.Raise -1, , "Registry root class not accepted."
                Exit Sub
            Case Else
                Err.Raise ClassRoot, , "Registry root class not supported."
                Exit Sub
        End Select
    End If
    
    p_sParam = p_sParam & "\" & m_sKey ' Build registry path
    
    ' Export via Windows® built-in Registry Editor
    ShellExecute 0&, "", "regedit", "/e """ & File & """ " & p_sParam, "", vbHide
End Sub

Public Function HasSubKey(Optional ClassRoot, Optional Key) As Boolean
Attribute HasSubKey.VB_Description = "Returns a value to determine if the registry path supplied contains subkeys."
'Returns a value to determine if the registry path supplied contains subkeys.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    Dim p_lKey As Long
    Dim p_lRet As Long
    
    If (m_bRemoteConnection) Then
        m_eClassRoot = m_lRemoteKey
    End If
    
    p_lRet = RegOpenKeyEx(m_eClassRoot, m_sKey, REG_OPTION_NON_VOLATILE, KEY_ENUMERATE_SUB_KEYS, p_lKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        Dim p_sBuffer As String * MAX_CLASS_NAME
        
        p_lRet = RegEnumKey(p_lKey, 0&, p_sBuffer, MAX_CLASS_NAME) ' Attempt to get subkey
        
        If (p_lRet = ERROR_SUCCESS) Then
            HasSubKey = True
        End If
        
        RegCloseKey p_lKey
    Else
        Err.Raise ClassRoot, , "Unable to open registry key '" & m_sKey & "' for enumerate subkeys access."
    End If
End Function

Public Sub Import(ByVal File As String)
Attribute Import.VB_Description = "Attempt to import a valid registry file to the registry."
'Attempt to import a valid registry file to the registry.
    File = Trim$(File)
    
    If (Len(Dir$(File))) Then
        ShellExecute 0&, "", "regedit", "/s """ & File & """", App.Path, vbHide
    Else
        Err.Raise 53 ' File not found
    End If
End Sub

Public Function KeyExists(Optional ClassRoot, Optional Key) As Boolean
Attribute KeyExists.VB_Description = "Returns/sets a value to determine if a key exists on the registry."
'Returns a value to determine if a key exists on the registry.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    Dim p_lKey As Long
    Dim p_lRet As Long
    
    If (m_bRemoteConnection) Then
        m_eClassRoot = m_lRemoteKey
    End If
    
    p_lRet = RegOpenKeyEx(m_eClassRoot, m_sKey, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, p_lKey)
    
    If (p_lRet = ERROR_SUCCESS) Or (p_lRet = ERROR_ACCESS_DENIED) Then
        KeyExists = True
        RegCloseKey p_lKey
    End If
End Function

Public Function RemoteConnect(Optional MachineName, Optional ClassRoot) As Boolean
Attribute RemoteConnect.VB_Description = "Establish a remote registry connection."
'Establish a remote registry connection.
    If (Not IsMissing(MachineName)) Then
        m_sMachineName = MachineName
    End If
    
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    Dim p_lKey As Long
    Dim p_lRet As Long
    
    p_lRet = RegConnectRegistry(m_sMachineName, m_eClassRoot, p_lKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        m_lRemoteKey = p_lKey
        m_eRemoteClassRoot = m_eClassRoot
        m_bRemoteConnection = True
        RemoteConnect = True
    Else
        Err.Raise ClassRoot, , "Unable to establish remote registry connection: '" & m_sMachineName & "'"
    End If
End Function

Public Function RemoteDisconnect() As Boolean
Attribute RemoteDisconnect.VB_Description = "Disconnect from remote registry connection."
'Disconnect from remote registry connection.
    If (m_bRemoteConnection) Then
        Dim p_lKey As Long
        Dim p_lRet As Long
        
        p_lRet = RegCloseKey(m_lRemoteKey)
        p_lRet = RegConnectRegistry("", m_eRemoteClassRoot, p_lKey) ' Connect from local system
        
        If (p_lRet = ERROR_SUCCESS) Then
            RegCloseKey p_lKey ' Then close registry handle
            
            m_lRemoteKey = 0
            m_eRemoteClassRoot = 0
            m_bRemoteConnection = False
            RemoteDisconnect = True
        Else
            Err.Raise m_eRemoteClassRoot, , "Unable to disconnect from remote registry connection."
        End If
    End If
End Function

Public Function ValueNameExists( _
        Optional ClassRoot, _
        Optional Key, _
        Optional ValueName) As Boolean
'Returns a value to determine if the specified valuename exists in a particular key.
    If (Not IsMissing(ClassRoot)) Then
        m_eClassRoot = ClassRoot
    End If
    
    If (Not IsMissing(Key)) Then
        m_sKey = Key
    End If
    
    If (Not IsMissing(ValueName)) Then
        m_sValueName = ValueName
    End If
    
    Dim p_lKey As Long
    Dim p_lRet As Long
    
    If (m_bRemoteConnection) Then
        m_eClassRoot = m_lRemoteKey
    End If
    
    p_lRet = RegOpenKeyEx(m_eClassRoot, m_sKey, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, p_lKey)
    
    If (p_lRet = ERROR_SUCCESS) Then
        p_lRet = RegQueryValueExString(p_lKey, m_sValueName, REG_OPTION_RESERVED, 0&, "", 0&)
        
        If (p_lRet = ERROR_SUCCESS) Or (p_lRet = ERROR_MORE_DATA) Then
            ValueNameExists = True
        End If
        
        RegCloseKey p_lKey
    Else
        Err.Raise ClassRoot, , "Unable to open registry key '" & m_sKey & "' for key query access."
    End If
End Function

' //-- Private procedures --//

Private Sub DeleteSubKeys(ByVal ClassRoot As String, Key As String)
' Deletes only subkeys with no child keys
    Dim p_lCtr          As Long
    Dim p_sKeyArray()   As String
    
    If (EnumerateKeys(p_sKeyArray, ClassRoot, Key)) Then
        For p_lCtr = LBound(p_sKeyArray) To UBound(p_sKeyArray)
            ' Attempt to delete every subkey found
            DeleteKey ClassRoot, Key & "\" & p_sKeyArray(p_lCtr)
        Next
    End If
End Sub

Private Function ExpandEnvironmentString(Value As String) As String
    Dim p_lLen      As Long
    Dim p_sBuffer   As String
    
    ' Get length of expanded string
    p_sBuffer = ""
    p_lLen = ExpandEnvironmentStrings(Value, p_sBuffer, p_lLen)
    
    ' Expand string
    p_sBuffer = String$(p_lLen, 0)
    p_lLen = ExpandEnvironmentStrings(Value, p_sBuffer, p_lLen)
    
    If (p_lLen > 0) Then
        ' Without the terminating null character
        ExpandEnvironmentString = Left$(p_sBuffer, p_lLen - 1)
    End If
End Function

Private Function GetParentKey(ByVal Value As String) As String
    Dim i As Long
    i = InStrRev(Value, "\")
    
    If (i > 0) Then
        GetParentKey = Left$(Value, i - 1) ' Get parent key
    End If
End Function

Private Function GetBigEndianValue(ByVal Value As Long) As Long
    CopyMemory ByVal VarPtr(GetBigEndianValue) + 3, Value, 1
    CopyMemory ByVal VarPtr(GetBigEndianValue) + 2, ByVal VarPtr(Value) + 1, 1
    CopyMemory ByVal VarPtr(GetBigEndianValue) + 1, ByVal VarPtr(Value) + 2, 1
    CopyMemory GetBigEndianValue, ByVal VarPtr(Value) + 3, 1
End Function

Private Sub ValidateMachineName(Value As String)
    If (Len(Value)) Then
        If (InStr(1, Value, "\\") = 0) Then
            Value = "\\" & Value ' This would be enough for now
        End If
    End If
End Sub

' //-- Class Procedure --//

Private Sub Class_Initialize()
    ' Unless these properties are set, these would be their default values
    m_eClassRoot = HKEY_CURRENT_USER
    m_eDataType = REG_SZ
    
    ' This is the required path format in writing software settings in the registry
    m_sKey = "Software\" & App.CompanyName & "\" & App.ProductName
End Sub

' Created by Noel A. Dacara | Copyright © 2003-2005 Davao City, Philippines
