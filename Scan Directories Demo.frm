VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Directories Demo"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
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
   ScaleHeight     =   7440
   ScaleWidth      =   9960
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "More Scanning Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6435
      Left            =   6720
      TabIndex        =   16
      Top             =   135
      Width           =   3060
      Begin VB.Frame Frame4 
         Caption         =   "Uncheck to see the difference"
         ForeColor       =   &H000000FF&
         Height          =   1455
         Left            =   300
         TabIndex        =   28
         Top             =   3075
         Width           =   2475
         Begin VB.CheckBox Check11 
            Caption         =   "Calculate CRC32 for Files"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   195
            TabIndex        =   31
            Top             =   975
            Value           =   1  'Checked
            Width           =   2220
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Show Files in Listview"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   195
            TabIndex        =   30
            Top             =   645
            Value           =   1  'Checked
            Width           =   2220
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Show History in Listview"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   195
            TabIndex        =   29
            Top             =   315
            Value           =   1  'Checked
            Width           =   2220
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Drive Type"
         Height          =   1995
         Left            =   300
         TabIndex        =   22
         Top             =   870
         Width           =   2475
         Begin VB.CheckBox Check8 
            Caption         =   "Include RAM-Disk Drives"
            Enabled         =   0   'False
            Height          =   255
            Left            =   135
            TabIndex        =   27
            Top             =   1575
            Width           =   2280
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Include Network Drives"
            Enabled         =   0   'False
            Height          =   255
            Left            =   135
            TabIndex        =   26
            Top             =   1260
            Width           =   2280
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Include CD-Rom Drives"
            Enabled         =   0   'False
            Height          =   255
            Left            =   135
            TabIndex        =   25
            Top             =   945
            Width           =   2280
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Include Removable Drives"
            Enabled         =   0   'False
            Height          =   255
            Left            =   135
            TabIndex        =   24
            Top             =   615
            Width           =   2280
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Include Hard Disk Drives"
            Enabled         =   0   'False
            Height          =   255
            Left            =   135
            TabIndex        =   23
            Top             =   285
            Value           =   1  'Checked
            Width           =   2280
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Scan All Drives by Type"
         Height          =   300
         Left            =   300
         TabIndex        =   21
         Top             =   375
         Width           =   2175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Current total size in bytes scanned:"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   5790
         Width           =   2565
      End
      Begin VB.Label NUMBYTES 
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   270
         TabIndex        =   32
         Top             =   6030
         Width           =   105
      End
      Begin VB.Label NUMFILES 
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   270
         TabIndex        =   20
         Top             =   5520
         Width           =   105
      End
      Begin VB.Label NUMFOLDERS 
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   270
         TabIndex        =   19
         Top             =   4995
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Current number of folders scanned:"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   4755
         Width           =   2580
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Current number of files scanned:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   5295
         Width           =   2370
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Abort"
      Enabled         =   0   'False
      Height          =   510
      Left            =   6720
      TabIndex        =   15
      Top             =   6750
      Width           =   3060
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5790
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Scan Directories Demo.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3765
      Left            =   180
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1575
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   6641
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Path/Filename (Right-click listview for more columns)"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "1100.97"
         Text            =   "CRC32"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "1000.06"
         Text            =   "Size (b)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "650.26"
         Text            =   "Attrib"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "2000.12"
         Text            =   "Last Modified"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "2000.12"
         Text            =   "Date Created"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "1200.18"
         Text            =   "Last Accessed"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scan Options:"
      Height          =   1140
      Left            =   180
      TabIndex        =   7
      Top             =   5445
      Width           =   6405
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   630
         Width           =   4560
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Custom Scan"
         Height          =   195
         Left            =   2610
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   270
         Width           =   1485
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   5760
         Style           =   1  'Simple Combo
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Combo2"
         Top             =   225
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Include Subdirectories"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   2000
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Custom Scan Path:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   675
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Scan Deepness:"
         Height          =   195
         Left            =   4500
         TabIndex        =   10
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Combo3"
      Top             =   1125
      Width           =   5190
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1395
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Combo2"
      Top             =   675
      Width           =   5190
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin Scanning"
      Default         =   -1  'True
      Height          =   510
      Left            =   180
      TabIndex        =   1
      Top             =   6750
      Width           =   6405
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1395
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   225
      Width           =   5190
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "File Attributes:"
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filter Settings:"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Directory Path:"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   270
      Width           =   1095
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPopups 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuPopups 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuPopups 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPopups 
         Caption         =   "Clear Listview"
         Index           =   3
      End
      Begin VB.Menu mnuPopups 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuPopups 
         Caption         =   "Columns"
         Index           =   5
         Begin VB.Menu mnuColumns 
            Caption         =   "Path\Filename"
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuColumns 
            Caption         =   "CRC32"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuColumns 
            Caption         =   "Size (b)"
            Index           =   2
         End
         Begin VB.Menu mnuColumns 
            Caption         =   "Attribute"
            Index           =   3
         End
         Begin VB.Menu mnuColumns 
            Caption         =   "Last Modified"
            Index           =   4
         End
         Begin VB.Menu mnuColumns 
            Caption         =   "Date Created"
            Index           =   5
         End
         Begin VB.Menu mnuColumns 
            Caption         =   "Last Accessed"
            Index           =   6
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Watch this area here--^ ...it contains all the form objects that supports events

Dim WithEvents SCANDIR As cScanDirectories
Attribute SCANDIR.VB_VarHelpID = -1
'Take note of the declaration above!!!

Private t As Single
Dim CRC32 As cCRC32

'##################################################################################

Private Sub SCANDIR_CurrentFile(File As String, Path As String, Delete As Boolean)
'   Tips: You can perform search and destroy operations here!
'         You can perform checksum checks and other file validation you want.
    Dim crc As String
    Dim l As ListItem
    
    If (Check10.Value = vbChecked) And (Check10.Enabled = True) Then
        Set l = ListView1.ListItems.Add(, , File)
            l.SubItems(2) = FormatNumber(SCANDIR.CurrentFileSize, 0, , , vbTrue)
            l.SubItems(3) = SCANDIR.CurrentFileAttribute
            l.SubItems(4) = SCANDIR.CurrentFileDate(EFT_LASTWRITETIME) & " " & SCANDIR.CurrentFileTime(EFT_LASTWRITETIME)
            l.SubItems(5) = SCANDIR.CurrentFileDate(EFT_CREATIONTIME) & " " & SCANDIR.CurrentFileTime(EFT_CREATIONTIME)
            l.SubItems(6) = SCANDIR.CurrentFileDate(EFT_LASTACCESSTIME) ' Last accessed do not have time included
            l.Selected = True
            l.EnsureVisible
        
        If (Check11.Value = vbChecked) Then
            crc = CRC32.FileChecksum(Path & "\" & File)
            l.SubItems(1) = crc
        End If
    End If
    
    NUMBYTES = NUMBYTES + SCANDIR.CurrentFileSize
    NUMFILES = SCANDIR.TotalFiles
End Sub

Private Sub SCANDIR_CurrentFolder(Path As String, Cancel As Boolean, Delete As Boolean)
'   Tips: You can perform search and destroy operations here!
'         Delete method here only supports for empty folders and without subfolders.
    
    'If Path = App.Path & "\New Folder" Then Delete = True
    
    Dim l As ListItem
    
    If (Check9.Value = vbChecked) Then
        Set l = ListView1.ListItems.Add(, , Mid$(Path, 4), , 1)
            l.SubItems(1) = "Drive " & Left$(Path, 2)
            l.SubItems(2) = FormatNumber(SCANDIR.CurrentFileSize, 0, , , vbTrue)
            l.SubItems(3) = SCANDIR.CurrentFileAttribute
            l.SubItems(4) = SCANDIR.CurrentFileDate(EFT_LASTWRITETIME) & " " & SCANDIR.CurrentFileTime(EFT_LASTWRITETIME)
            l.SubItems(5) = SCANDIR.CurrentFileDate(EFT_CREATIONTIME) & " " & SCANDIR.CurrentFileTime(EFT_CREATIONTIME)
            l.SubItems(6) = SCANDIR.CurrentFileDate(EFT_LASTACCESSTIME) ' Last accessed do not have time included
            l.Selected = True
            l.ForeColor = vbBlue
            l.ListSubItems(1).ForeColor = RGB(0, 150, 0)
            l.ToolTipText = ListView1.SelectedItem.Text
            l.EnsureVisible
    End If
    
    NUMFOLDERS = SCANDIR.TotalFolders
End Sub

Private Sub SCANDIR_DoneScanning(TotalFolders As Long, TotalFiles As Long)
    Command1.Caption = "Begin Scanning"
    Command2.Enabled = False
    
    MsgBox "Total Folders Scanned: " & TotalFolders & vbNewLine & _
           "Total Files Scanned  : " & TotalFiles & vbNewLine & vbNewLine & _
           "Total Size in Scanned: " & FormatNumber(NUMBYTES / 1024 / 1024, 2) & " MB" & vbNewLine & vbNewLine & _
           "Total Scan Time: " & CDbl(Timer - t) & " seconds.", vbInformation, "Done Scanning"
End Sub

'##################################################################################

'Below are the usual form procedures

Private Sub Form_Load()
    Set SCANDIR = New cScanDirectories
    Set CRC32 = New cCRC32
    
    'You can also search the current drive simply by leaving the start path empty
    'Or a whole drive by specifying only the drive letter...
    Combo1.Text = "C:\Program Files"
    
    With Combo2
        .AddItem "*.com"
        .AddItem "*.exe"
        .AddItem "*.dll"
        .AddItem "*.ocx"
        .AddItem "*.scr"
        .AddItem "*.vbs"
        .AddItem "*.com|*.exe|*.dll|*.ocx|*.scr|*.vbs"
        .AddItem "*.*" 'All files with extension
        .AddItem "*" 'All files (includes files w/o extension)
        .ListIndex = .ListCount - 1
    End With
    
    Combo3.Text = "Default to search for all files (can be changed)"
    Combo4.Text = 0
    
    With Combo5
        .AddItem "SCAN_ALLDESKTOP"
        .ItemData(.NewIndex) = ECP_COMMON_DESKTOPDIRECTORY
        .AddItem "SCAN_ALLSTARTMENU"
        .ItemData(.NewIndex) = ECP_COMMON_STARTMENU
        .AddItem "SCAN_ALLSTARTUP"
        .ItemData(.NewIndex) = ECP_COMMON_STARTUP
        .AddItem "SCAN_FONTS"
        .ItemData(.NewIndex) = ECP_FONTS
        .AddItem "SCAN_PROGRAMFILES"
        .ItemData(.NewIndex) = ECP_PROGRAM_FILES
        .AddItem "SCAN_SYSTEMDIR"
        .ItemData(.NewIndex) = ECP_SYSTEM
        .AddItem "SCAN_TEMPFILESDIR"
        .ItemData(.NewIndex) = ECP_TEMPORARYFILES
        .AddItem "SCAN_USERDESKTOP"
        .ItemData(.NewIndex) = ECP_DESKTOP
        .AddItem "SCAN_USERDOCUMENTS"
        .ItemData(.NewIndex) = ECP_PERSONAL
        .AddItem "SCAN_USERRECENTS"
        .ItemData(.NewIndex) = ECP_RECENT
        .AddItem "SCAN_USERSTARTMENU"
        .ItemData(.NewIndex) = ECP_STARTMENU
        .AddItem "SCAN_USERSTARTUP"
        .ItemData(.NewIndex) = ECP_STARTUP
        .AddItem "SCAN_WINDOWSDIR"
        .ItemData(.NewIndex) = ECP_WINDOWS
        
        .ListIndex = 0
    End With
    
    Check1.Value = vbChecked
    Check2.Value = vbChecked
    Check2.Value = vbUnchecked 'To trigger the click event
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Begin Scanning" Then
        ListView1.ListItems.Clear
        NUMBYTES = 0
        
        t = Timer
        SCANDIR.ScanDrives = (Check3.Value = vbChecked)
        
        If Check3.Value = vbChecked Then
            SCANDIR.ScanDriveType = 0
            If Check4.Value = vbChecked Then
                SCANDIR.ScanDriveType = SCANDIR.ScanDriveType Or EDT_FIXED
            End If
            If Check5.Value = vbChecked Then
                SCANDIR.ScanDriveType = SCANDIR.ScanDriveType Or EDT_REMOVABLE
            End If
            If Check6.Value = vbChecked Then
                SCANDIR.ScanDriveType = SCANDIR.ScanDriveType Or EDT_CDROM
            End If
            If Check7.Value = vbChecked Then
                SCANDIR.ScanDriveType = SCANDIR.ScanDriveType Or EDT_REMOTE
            End If
            If Check8.Value = vbChecked Then
                SCANDIR.ScanDriveType = SCANDIR.ScanDriveType Or EDT_RAMDISK
            End If
        End If
        SCANDIR.StartPath = Combo1.Text
        SCANDIR.Filter = Combo2.Text
        SCANDIR.SubDirectories = (Check1.Value = vbChecked)
        SCANDIR.ScanDeep = Combo4.Text
        SCANDIR.CustomScan = (Check2.Value = vbChecked)
        SCANDIR.CustomScanPath = Combo5.ItemData(Combo5.ListIndex)
        
        Command1.Caption = "Pause Scanning"
        Command2.Enabled = True
        
        SCANDIR.BeginScanning
    ElseIf Command1.Caption = "Pause Scanning" Then
        Command1.Caption = "Resume Scanning"
        SCANDIR.PauseScanning
    Else
        Command1.Caption = "Pause Scanning"
        SCANDIR.ResumeScanning
    End If
End Sub

Private Sub Command2_Click()
    SCANDIR.CancelScanning
    Command1.Caption = "Begin Scanning"
    Command2.Enabled = False
End Sub

Private Sub Check1_Click()
    Combo4.Locked = Not (Check1.Value = vbChecked)
    If (Combo4.Locked = True) Then
        Combo4.BackColor = vbButtonFace
    Else
        Combo4.BackColor = vbWhite
    End If
End Sub

Private Sub Check2_Click()
    If (Check2.Value = vbChecked) Then
        Combo1.Locked = True
        Combo1.BackColor = vbButtonFace
        Combo5.Locked = False
        Combo5.BackColor = vbWindowBackground
    Else
        Combo1.Locked = False
        Combo1.BackColor = vbWindowBackground
        Combo5.Locked = True
        Combo5.BackColor = vbButtonFace
    End If
End Sub

Private Sub Check3_Click()
    Check4.Enabled = (Check3.Value = vbChecked)
    Check5.Enabled = (Check3.Value = vbChecked)
    Check6.Enabled = (Check3.Value = vbChecked)
    Check7.Enabled = (Check3.Value = vbChecked)
    Check8.Enabled = (Check3.Value = vbChecked)
    
    If (Check3.Value = vbChecked) Then
        Combo1.Locked = True
        Combo1.BackColor = vbButtonFace
        Check2.Enabled = False
        Combo5.Locked = True
        Combo5.BackColor = vbButtonFace
    Else
        Check2.Enabled = True
        Check2_Click
    End If
End Sub

Private Sub Check9_Click()
    Check10.Enabled = (Check9.Value = vbChecked)
    If (Check10.Value = vbChecked) Then
        Check11.Enabled = (Check9.Value = vbChecked)
    Else
        ' Just do nothing
    End If
End Sub

Private Sub Check10_Click()
    Check11.Enabled = (Check10.Value = vbChecked)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then
        SCANDIR.CancelScanning
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SCANDIR = Nothing
    Set CRC32 = Nothing
End Sub

Private Sub mnuColumns_Click(Index As Integer)
    mnuColumns(Index).Checked = Not mnuColumns(Index).Checked
    If (mnuColumns(Index).Checked = True) Then
        ListView1.ColumnHeaders(Index + 1).Width = ListView1.ColumnHeaders(Index + 1).Tag
    Else
        ListView1.ColumnHeaders(Index + 1).Width = 0
    End If
End Sub

Private Sub mnuPopups_Click(Index As Integer)
    Select Case Index
        Case 0: Command1_Click
        Case 1: Command2_Click
        Case 3: ListView1.ListItems.Clear
    End Select
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = vbRightButton) Then
        mnuPopups(0).Caption = Command1.Caption
        mnuPopups(1).Caption = Command2.Caption
        mnuPopups(1).Enabled = Command2.Enabled
        mnuPopups(3).Enabled = (ListView1.ListItems.Count > 0)
        Dim i As Integer
        For i = 1 To 6
            mnuColumns(i).Checked = (ListView1.ColumnHeaders(i + 1).Width <> 0)
        Next i
        PopupMenu mnuPopup, , , , mnuPopups(0)
    End If
End Sub

