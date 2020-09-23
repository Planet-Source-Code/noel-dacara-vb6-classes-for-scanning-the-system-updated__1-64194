VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Processes Demo"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8460
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7395
      Top             =   4005
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Refresh List of Processes"
      Default         =   -1  'True
      Height          =   465
      Left            =   225
      TabIndex        =   0
      Top             =   5760
      Width           =   8040
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scan Options:"
      Height          =   555
      Left            =   225
      TabIndex        =   3
      Top             =   5025
      Width           =   5805
      Begin VB.CheckBox Check2 
         Caption         =   "Show DLLs"
         Height          =   240
         Left            =   2415
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show System Processes"
         Height          =   240
         Left            =   180
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   2040
      End
      Begin VB.Label NUMPROC 
         AutoSize        =   -1  'True
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   6
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number of Processes:"
         Height          =   195
         Left            =   3735
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Terminate Process"
      Height          =   465
      Left            =   6180
      TabIndex        =   1
      Top             =   5115
      Width           =   2085
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4680
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   225
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   8255
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Base Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Process Path"
         Object.Width           =   7426
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Process ID"
         Object.Width           =   1685
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Modules"
         Object.Width           =   1685
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Watch this area here--^ ...it contains all the form objects that supports events

Dim WithEvents SCANPROC As cScanProcesses
Attribute SCANPROC.VB_VarHelpID = -1
'Take note of the declaration above!!!

Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

Dim m_sProcess  As String
Dim m_sTime     As Single
Dim FILEICON    As cGetFileIcon

'##################################################################################

Private Sub SCANPROC_CurrentModule(Process As String, ID As Long, Module As String, File As String)
'   Tips: You can perform checksum checks here indiviually for each file...
    Dim lsv As ListItem
    
    Set lsv = ListView1.ListItems.Add(, , Module)
    
    With lsv
        .ForeColor = RGB(0, 150, 0) ' Dark green
        .SubItems(1) = File
        .ListSubItems(1).ToolTipText = File
'        .Selected = True
'        .EnsureVisible
    End With
End Sub

Private Sub SCANPROC_CurrentProcess(Name As String, File As String, ID As Long, Modules As Long)
'   Tips: You can perform checksum checks here indiviually for each file...
    Dim p_HasImage As Boolean
    
    If (File <> "SYSTEM") Then
        On Error Resume Next
        ImageList1.ListImages(Name).Tag = ""   'Just to test if this item exists
        
        If (Err.Number <> 0) Then
            Err.Clear
            ImageList1.ListImages.Add , Name, FILEICON.Icon(File, SmallIcon)
            p_HasImage = (Err.Number = 0)
        Else
            p_HasImage = True
        End If
    End If
    
    Dim lsv As ListItem
    
    If (p_HasImage = True) Then
        Set lsv = ListView1.ListItems.Add(, "#" & Name & ID, Name, , Name)
    Else
        Set lsv = ListView1.ListItems.Add(, "#" & Name & ID, Name)
    End If
    
    With lsv
        .ForeColor = vbBlue
        .SubItems(1) = File
        .SubItems(2) = ID
        .SubItems(3) = Modules
        .ListSubItems(2).ForeColor = vbRed
        .ListSubItems(1).ToolTipText = File
'        .Selected = True
'        .EnsureVisible
    End With
    
    If (m_sProcess <> "#" & Name & ID) Then
        Modules = 0
    End If
End Sub

Private Sub SCANPROC_DoneScanning(TotalProcess As Long)
    Dim p_Elapsed As Single
    p_Elapsed = Timer - m_sTime
    
    LockWindowUpdate 0& ' Enable listview repaint
    
    'Debug.Print "Total Number of Process Detected: " & TotalProcess & vbNewLine & "Total Scan Time: " & p_Elapsed & vbNewLine
    NUMPROC = TotalProcess
End Sub

'##################################################################################

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If (Check2.Value <> vbChecked) Then
        Exit Sub ' What for?
    End If
    
    Dim i As Long
    i = Item.Index
    
    If (m_sProcess = Item.Key) Then
        m_sProcess = ""
    Else
        m_sProcess = Item.Key
    End If
    
    Call Check1_Click
    
    On Error Resume Next
    ListView1.ListItems(i).Selected = True
    ListView1.SelectedItem.EnsureVisible
End Sub

Private Sub Check1_Click()
    ListView1.ListItems.Clear
    
    SCANPROC.SystemProcesses = (Check1.Value = vbChecked)
    SCANPROC.ProcessModules = (Check2.Value = vbChecked)
    
    m_sTime = Timer
    
    LockWindowUpdate ListView1.hWnd ' Prevent listview repaints
    SCANPROC.BeginScanning
End Sub

Private Sub Check2_Click()
    If (Check2.Value = vbChecked) Then
        Command2.Caption = "&Refresh List of Processes (Select a process on list to show modules used)"
    Else
        Command2.Caption = "&Refresh List of Processes"
    End If
End Sub

Private Sub Command1_Click()
    If MsgBox("Are you sure to terminate this process?", vbExclamation + vbYesNoCancel + vbDefaultButton2, "Terminate Process") = vbYes Then
        If (SCANPROC.TerminateProcess(ListView1.SelectedItem.SubItems(2)) = True) Then
            Call Check1_Click
        End If
    End If
End Sub

Private Sub Command2_Click()
    Call Check1_Click
End Sub

Private Sub Form_Initialize()
    Set SCANPROC = New cScanProcesses
    Set FILEICON = New cGetFileIcon
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then
        SCANPROC.CancelScanning
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call Check1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SCANPROC = Nothing
    Set FILEICON = Nothing
End Sub

