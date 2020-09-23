VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScanNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --------------------------------------------------------------------------------
' Name:     cScanNetwork class
' Author:   Noel A. Dacara (noeldacara@hotmail.com)
' Date:     11 January 2006
'
' You can freely use and distribute this file with or without permission provided
' that credits herein are kept unmodified and having acknowledgements is observed.
'
' Credits:  Powersoft Programming at http://psprogramming.virtualave.net/
'
' Tips:     To attain the maximum scanning speed of your machine with this class,
'           you may need to close unnecessary programs running in your computer..
'
' Inquiries or comments regarding my works and whatsoever are greatly appreciated.
' --------------------------------------------------------------------------------

' API Declarations
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetQueueStatus Lib "user32.dll" (ByVal fuFlags As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32.dll" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, ByRef lpcCount As Long, ByRef lpBuffer As Any, ByRef lpBufferSize As Long) As Long
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (ByRef lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long

' Modified API Declarations
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, ByRef lpNetResource As Any, ByRef lphEnum As Long) As Long

' API Constants
Private Const ERROR_EXTENDED_ERROR      As Long = 1208&
Private Const ERROR_INVALID_PARAMETER   As Long = 87
Private Const ERROR_MORE_DATA           As Long = 234
Private Const ERROR_NO_NETWORK          As Long = 1222&
Private Const GMEM_FIXED                As Long = &H0
Private Const GMEM_ZEROINIT             As Long = &H40
Private Const GPTR                      As Long = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const INFINITE                  As Long = &HFFFFFFFF
Private Const MAX_DESTINATION_BUFFER    As Long = &H1000
Private Const MAX_RESOURCES             As Long = 256
Private Const NO_ERROR                  As Long = 0
Private Const RESOURCE_CONNECTED        As Long = &H1
Private Const RESOURCE_GLOBALNET        As Long = &H2
Private Const RESOURCETYPE_ANY          As Long = &H0
Private Const RESOURCETYPE_DISK         As Long = &H1
Private Const RESOURCEUSAGE_ALL         As Long = &H0

' Resource Display Type Contants
Private Const RESOURCEDISPLAYTYPE_DIRECTORY     As Long = &H9
Private Const RESOURCEDISPLAYTYPE_DOMAIN        As Long = &H1
Private Const RESOURCEDISPLAYTYPE_FILE          As Long = &H4
Private Const RESOURCEDISPLAYTYPE_GENERIC       As Long = &H0
Private Const RESOURCEDISPLAYTYPE_GROUP         As Long = &H5
Private Const RESOURCEDISPLAYTYPE_NDSCONTAINER  As Long = &HB
Private Const RESOURCEDISPLAYTYPE_NETWORK       As Long = &H6
Private Const RESOURCEDISPLAYTYPE_ROOT          As Long = &H7
Private Const RESOURCEDISPLAYTYPE_SERVER        As Long = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE         As Long = &H3
Private Const RESOURCEDISPLAYTYPE_SHAREADMIN    As Long = &H8
Private Const RESOURCEDISPLAYTYPE_TREE          As Long = &HA

'GetQueueStatus Flag
Private Const QS_ALLEVENTS As Long = &HBF

' API Types
Private Type NETRESOURCE
    dwScope         As Long
    dwType          As Long
    dwDisplayType   As Long
    dwUsage         As Long
    lpLocalName     As String
    lpRemoteName    As String
    lpComment       As String
    lpProvider      As String
End Type

' Class Variables
Private m_bCancel           As Boolean
Private m_bPaused           As Boolean
Private m_bScanComputers    As Boolean
Private m_bScanning         As Boolean
Private m_lTotalDomains     As Long
Private m_lTotalComputers   As Long
Private m_sDomain           As String

' //-- Class Events --//

Public Event CurrentComputer(Name As String, Domain As String)
Attribute CurrentComputer.VB_Description = "Occurs everytime a computer is scanned from a domain."
'Occurs everytime a computer is scanned from a domain.
Public Event CurrentDomain(Name As String, Provider As String)
Attribute CurrentDomain.VB_Description = "Occurs everytime a domain is scanned."
'Occurs everytime a domain is scanned.
Public Event DoneScanning(TotalDomains As Long, TotalComputers As Long)
Attribute DoneScanning.VB_Description = "Occurs after scanning has completed or cancelled."
'Occurs after scanning has completed or cancelled.

' //-- Properties --//

Public Property Get Domain() As String
Attribute Domain.VB_Description = "Returns/sets the domain to scan for computers. Empty to scan all computers within the network."
'Returns/sets the domain to scan for computers. Empty to scan all computers within the network.
    Domain = m_sDomain
End Property

Public Property Let Domain(Value As String)
    m_sDomain = Value
End Property

Public Property Get Paused() As Boolean
Attribute Paused.VB_Description = "Returns True if scanning process has been temporarily stopped."
'Returns True if scanning process has been temporarily stopped.
    Paused = m_bPaused
End Property

Public Property Get ScanComputers() As Boolean
Attribute ScanComputers.VB_Description = "Returns/sets a value whether to include scanning for computers or just the domains."
'Returns/sets a value whether to include scanning for computers or just the domains.
    ScanComputers = m_bScanComputers
End Property

Public Property Let ScanComputers(Value As Boolean)
    ' This property is ignored if Domain is set
    m_bScanComputers = Value
End Property

Public Property Get Scanning() As Boolean
Attribute Scanning.VB_Description = "Returns True if class is currently on scanning state."
'Returns True if class is currently on scanning state.
    Scanning = m_bScanning
End Property

Public Property Get TotalDomains() As Long
Attribute TotalDomains.VB_Description = "Returns the current number of domains scanned."
'Returns the current number of domains scanned.
    TotalDomains = m_lTotalDomains
End Property

Public Property Get TotalComputers() As Long
Attribute TotalComputers.VB_Description = "Returns the current number of computers scanned."
'Returns the current number of computers scanned.
    TotalComputers = m_lTotalComputers
End Property

' //-- Procedures --//

Public Sub BeginScanning()
Attribute BeginScanning.VB_Description = "Start scanning for domains and computers within the network."
'Start scanning for domains or computers within the network.
    If (m_bPaused) Then
        ResumeScanning ' Resume scanning instead
        Exit Sub
    End If
    
    If (m_bScanning) Then
        Exit Sub 'Avoid cascading scans
    End If
    
    m_lTotalDomains = 0
    m_lTotalComputers = 0
    
    m_bCancel = False
    m_bScanning = True
    
    If (Len(m_sDomain) = 0) Then
        ScanForDomains ' Scan all computers on the network
    Else
        ScanForComputers m_sDomain
    End If
    
    m_bScanning = False
    m_bCancel = False
    
    RaiseEvent DoneScanning(m_lTotalDomains, m_lTotalComputers)
End Sub

Public Sub CancelScanning()
Attribute CancelScanning.VB_Description = "Abort the whole scanning process."
'Abort the whole scanning process.
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

Private Sub ScanForComputers(Domain As String)
    Dim i               As Long
    Dim p_lBufferSize   As Long
    Dim p_lBufferPtr    As Long
    Dim p_lCount        As Long
    Dim p_lEnum         As Long
    Dim p_lRet          As Long
    Dim p_lStructOffset As Long
    Dim p_sComputer     As String
    Dim p_tResource     As NETRESOURCE
    
    p_tResource.lpRemoteName = Domain
    p_tResource.dwDisplayType = RESOURCEDISPLAYTYPE_DOMAIN
    
    If (WNetOpenEnum(RESOURCE_GLOBALNET, _
                     RESOURCETYPE_DISK, _
                     RESOURCEUSAGE_ALL, _
                     p_tResource, _
                     p_lEnum) = NO_ERROR) Then
                     
        p_lBufferSize = MAX_DESTINATION_BUFFER
        p_lBufferPtr = GlobalAlloc(GPTR, p_lBufferSize) ' Allocate memory
        
        Do
            p_lCount = INFINITE
            p_lRet = WNetEnumResource(p_lEnum, p_lCount, p_lBufferPtr, p_lBufferSize)
            
            If (p_lRet = NO_ERROR) Then
                ' Offset to the first NETRESOURCE structure
                p_lStructOffset = p_lBufferPtr
                
                For i = 1 To p_lCount
                    If (GetQueueStatus(QS_ALLEVENTS)) Then
                        DoEvents ' Process system events
                    End If
                    
                    If (m_bPaused) Then
                        m_bScanning = False
                        While (m_bPaused)
                            DoEvents ' Don't process next instruction until resumed
                        Wend
                        m_bScanning = True
                    End If
                    
                    If (m_bCancel) Then
                        Exit For
                    End If
                    
                    ' Get NETRESOURCE structure from memory
                    CopyMemory p_tResource, ByVal p_lRet, LenB(p_tResource)
                    
                    p_sComputer = p_tResource.lpRemoteName
                    p_sComputer = Left$(p_sComputer, lstrlen(p_sComputer))
                    
                    If (Len(p_sComputer)) Then
                        p_lRet = InStr(1, p_sComputer, "\\")
                        
                        If (p_lRet > 0) Then
                            p_sComputer = Mid$(p_sComputer, p_lRet + 2)
                        End If
                        
                        m_lTotalComputers = m_lTotalComputers + 1
                        RaiseEvent CurrentComputer(p_sComputer, Domain)
                    End If
                    
                    ' Set offset to the next NETRESOURCE structure
                    p_lStructOffset = p_lStructOffset + LenB(p_tResource)
                Next
            ElseIf (p_lRet = ERROR_MORE_DATA) Then
                ' p_lBufferPtr is not large enough to handle all of the
                ' information in the NETRESOURCE structure, p_lBufferSize
                ' has been update to hold the required amount of space
                
                ' Resize memory
                p_lBufferPtr = GlobalReAlloc(p_lBufferPtr, p_lBufferSize, GPTR)
            Else
                RaiseNetworkError Err.LastDllError
                Exit Do
            End If
        Loop While p_lCount > 0
        
        GlobalFree p_lBufferPtr ' Free allocated memory
        WNetCloseEnum p_lEnum ' Close network resource enumeration
    Else
        RaiseNetworkError Err.LastDllError
    End If
End Sub

Private Sub ScanForDomains()
    Dim i               As Long
    Dim p_lBufferSize   As Long
    Dim p_lCount        As Long
    Dim p_lEnum         As Long
    Dim p_sDomainName   As String
    Dim p_sProvider     As String
    Dim p_tResources()  As NETRESOURCE
    
    If (WNetOpenEnum(RESOURCE_GLOBALNET, _
                     RESOURCETYPE_ANY, _
                     RESOURCEUSAGE_ALL, _
                     ByVal 0&, _
                     p_lEnum) = NO_ERROR) Then
                     
        ReDim p_tResources(1 To MAX_RESOURCES) As NETRESOURCE
        
        p_lBufferSize = UBound(p_tResources) * Len(p_tResources(0))
        p_lCount = INFINITE ' Return as many entries as possible
        
        If (WNetEnumResource(p_lEnum, _
                             p_lCount, _
                             p_tResources(1), _
                             p_lBufferSize) = NO_ERROR) Then
                             
            For i = 1 To p_lCount ' Iterate through all domains
                If (GetQueueStatus(QS_ALLEVENTS)) Then
                    DoEvents ' Process other system events other than this function
                End If
                
                If (m_bPaused) Then
                    m_bScanning = False
                    While (m_bPaused)
                        DoEvents ' Don't process next instruction until resumed
                    Wend
                    m_bScanning = True
                End If
                
                If (m_bCancel) Then
                    Exit For ' Abort scanning
                End If
                
                p_sDomainName = p_tResources(1).lpRemoteName
                p_sDomainName = Left$(p_sDomainName, lstrlen(p_sDomainName))
                
                p_sProvider = p_tResources(1).lpProvider
                p_sProvider = Left$(p_sProvider, lstrlen(p_sProvider))
                
                If (Len(p_sDomainName)) Then
                    m_lTotalDomains = m_lTotalDomains + 1
                    RaiseEvent CurrentDomain(p_sDomainName, p_sProvider)
                    
                    If (m_bScanComputers) Then
                        ' Scan domain for computers
                        ScanForComputers p_sDomainName
                    End If
                End If
            Next
        Else
            RaiseNetworkError Err.LastDllError
        End If
        
        WNetCloseEnum p_lEnum ' Close network resource enumeration
    Else
        RaiseNetworkError Err.LastDllError
    End If
End Sub

Private Sub RaiseNetworkError(nError As Long)
    Select Case nError
        Case ERROR_NO_NETWORK
            Err.Raise nError, , "No network is present."
        Case ERROR_EXTENDED_ERROR
            Dim p_lError As Long
            Dim p_sDesc As String
            Dim p_sSource As String
            
            Const MAXERRORLENGTH As Long = 256 ' 128 formerly
            
            p_sDesc = String$(MAXERRORLENGTH, 0)
            p_sSource = String$(MAXERRORLENGTH, 0)
            
            ' Get network error description
            If (WNetGetLastError(p_lError, _
                                 p_sSource, _
                                 MAXERRORLENGTH, _
                                 p_sDesc, _
                                 MAXERRORLENGTH) = NO_ERROR) Then
                                 
                p_sDesc = Left$(p_sDesc, lstrlen(p_sDesc))
                p_sSource = Left$(p_sSource, lstrlen(p_sSource))
                
                ' Raise an error
                Err.Raise p_lError, p_sSource, p_sDesc
            End If
        Case ERROR_INVALID_PARAMETER
            Err.Raise nError, , "Unable to successfully enumerate domain computers."
        Case Else
            Debug.Print "An unrecognized error has occured with error number "; nError
    End Select
End Sub

' //-- Class Procedures --//

Private Sub Class_Initialize()
    ' Default property values if not set by user
    m_bScanComputers = True
End Sub

Private Sub Class_Terminate()
    If (m_bScanning Or m_bPaused) Then
        CancelScanning 'Just to ensure that everything will be alright
    End If
End Sub

' Created by Noel A. Dacara | Copyright © 2003-2005 Davao City, Philippines
