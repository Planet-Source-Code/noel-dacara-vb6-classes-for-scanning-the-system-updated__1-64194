VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Registry Demo"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10320
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Abort"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6885
      TabIndex        =   25
      Top             =   7335
      Width           =   3225
   End
   Begin VB.Frame Frame2 
      Caption         =   "More Scanning Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7065
      Left            =   6885
      TabIndex        =   19
      Top             =   120
      Width           =   3240
      Begin VB.CheckBox Check8 
         Caption         =   "Perform Full Registry Scan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   330
         TabIndex        =   26
         Top             =   495
         Width           =   2595
      End
      Begin VB.Frame Frame3 
         Caption         =   "Uncheck to see the difference"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2265
         Left            =   240
         TabIndex        =   24
         Top             =   1740
         Width           =   2745
         Begin VB.CheckBox Check7 
            Caption         =   "Show Building of Lists"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   225
            TabIndex        =   32
            Top             =   1770
            Value           =   1  'Checked
            Width           =   2325
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Put Listview Data Tooltips"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   225
            TabIndex        =   30
            Top             =   1425
            Value           =   1  'Checked
            Width           =   2325
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Get Value of Registry Data"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   225
            TabIndex        =   29
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2325
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Show Registry Key Data"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   225
            TabIndex        =   28
            Top             =   735
            Value           =   1  'Checked
            Width           =   2325
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Show History in Listview"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   225
            TabIndex        =   27
            Top             =   390
            Value           =   1  'Checked
            Width           =   2325
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Note: This option may take a longer span of time inorder to complete of course depending on size of registry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   660
         Left            =   315
         TabIndex        =   31
         Top             =   930
         Width           =   2625
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   735
         Picture         =   "Scan Registry Demo.frx":0000
         Top             =   4515
         Width           =   1830
      End
      Begin VB.Label NUMKEYS 
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
         Left            =   420
         TabIndex        =   23
         Top             =   6150
         Width           =   105
      End
      Begin VB.Label NUMDATA 
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
         Left            =   420
         TabIndex        =   22
         Top             =   6645
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Current number of Keys Scanned:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   21
         Top             =   5895
         Width           =   2445
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Current number of Data Scanned:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   6420
         Width           =   2445
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1980
      ScaleHeight     =   855
      ScaleWidth      =   2805
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3105
      Visible         =   0   'False
      Width           =   2805
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   2475
         TabIndex        =   18
         Top             =   345
         Width           =   105
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Building List..."
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
         Left            =   240
         TabIndex        =   17
         Top             =   345
         Width           =   1140
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5895
      Top             =   5115
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
            Picture         =   "Scan Registry Demo.frx":11C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4350
      Left            =   180
      TabIndex        =   15
      Top             =   1530
      Width           =   6500
      _ExtentX        =   11456
      _ExtentY        =   7673
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   7373
      EndProperty
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4410
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Text            =   "Combo6"
      Top             =   1050
      Width           =   2250
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Text            =   "Combo5"
      Top             =   1050
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scan Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   180
      TabIndex        =   10
      Top             =   5985
      Width           =   6495
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   675
         Width           =   4545
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Custom Scanning"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2460
         TabIndex        =   6
         Top             =   300
         Width           =   1590
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5820
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Text            =   "Combo3"
         Top             =   255
         Width           =   420
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Include Subkeys"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Custom Scan Path:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Scan Deepness:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4515
         TabIndex        =   11
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin Scanning"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   180
      TabIndex        =   0
      Top             =   7335
      Width           =   6480
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   630
      Width           =   5355
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   225
      Width           =   6480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Filter Keys:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3495
      TabIndex        =   14
      Top             =   1095
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Filter Data:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   13
      Top             =   1095
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registry Path:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   675
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Watch this area here--^ ...it contains all the form objects that supports events

Dim WithEvents SCANREG As cScanRegistry
Attribute SCANREG.VB_VarHelpID = -1
'Take note of the declaration above!!!

Private t As Single
Dim REG As cAdvanceRegistry

'##################################################################################

Private Sub SCANREG_BuildingDataList(Index As Long, Total As Long)
    If Check7.Value = vbChecked Then
        If Total > 1000 Then
            Picture1.Visible = True
            Label8 = "Building Data List"
            Label9 = Index
        End If
        If Index = Total Then
            Picture1.Visible = False
        End If
    End If
End Sub

Private Sub SCANREG_BuildingKeyList(Index As Long, Total As Long)
    If Check7.Value = vbChecked Then
        If Total > 1000 Then
            Picture1.Visible = True
            Label8 = "Building Key List"
            Label9 = Index
        End If
        If Index = Total Then
            Picture1.Visible = False
        End If
    End If
End Sub

Private Sub SCANREG_CurrentData(Value As String, Key As String, Root As EScanRegistryRoots, Delete As Boolean)
    ' Tips: You can perform search and destroy operations here.
    
    ' Debug.Print Key & "\" & Value
    If Check3.Value = vbChecked And Check4.Value = vbChecked Then
        If Len(Value) = 0 Then
            ' Default Key Data doesn't have value names (Empty) or ""
            ' so let's make it as (Default) here for clarification...
            ListView1.ListItems.Add , , "(Default)"
        Else
            ListView1.ListItems.Add , , Value
        End If
        
        On Error Resume Next
        ListView1.ListItems(ListView1.ListItems.Count).Selected = True
        If Check5.Value = vbChecked Then
           ListView1.SelectedItem.SubItems(1) = REG.ValueEx(Root, Key, Value)
        End If
        If Check6.Value = vbChecked Then
            ListView1.SelectedItem.ListSubItems(1).ToolTipText = ListView1.SelectedItem.SubItems(1)
        End If
        ListView1.SelectedItem.EnsureVisible
    End If
    
    NUMDATA = SCANREG.TotalData
End Sub

Private Sub SCANREG_CurrentKey(Key As String, Root As EScanRegistryRoots, Delete As Boolean)
    ' Tips: You can perform search and destroy operations here.
    If Check3.Value = vbChecked Then
        Dim r As String
        Select Case Root
            Case SCAN_CLASSES_ROOT
                r = "CLASSES_ROOT"
            Case SCAN_CURRENT_USER
                r = "CURRENT_USER"
            Case SCAN_LOCAL_MACHINE
                r = "LOCAL_MACHINE"
            Case SCAN_USERS
                r = "USERS"
        End Select
        
        ListView1.ListItems.Add(, , r, , 1).SubItems(1) = Key
        
        On Error Resume Next
        ListView1.ListItems(ListView1.ListItems.Count).Selected = True
        ListView1.SelectedItem.ForeColor = vbBlue
        ListView1.SelectedItem.ListSubItems(1).ForeColor = vbRed
        If Check6.Value = vbChecked Then
            ListView1.SelectedItem.ListSubItems(1).ToolTipText = ListView1.SelectedItem.SubItems(1)
        End If
        ListView1.SelectedItem.EnsureVisible
    End If
    
    NUMKEYS = SCANREG.TotalKeys
End Sub

Private Sub SCANREG_DoneScanning(TotalData As Long, TotalKeys As Long)
    Command1.Caption = "Begin Scanning"
    Command2.Enabled = False
    Picture1.Visible = False
    
    MsgBox "Total Keys Scanned: " & TotalKeys & vbNewLine & _
           "Total Data Scanned: " & TotalData & vbNewLine & vbNewLine & _
           "Total Scan Time: " & CDbl(Timer - t) & " seconds.", vbInformation, "Done Scanning"
End Sub

'##################################################################################

Private Sub Form_Load()
    Set SCANREG = New cScanRegistry
    Set REG = New cAdvanceRegistry
    
    Combo1.AddItem "SCAN_CLASSES_ROOT", 0
    Combo1.AddItem "SCAN_CURRENT_USER", 1
    Combo1.AddItem "SCAN_LOCAL_MACHINE", 2

    Combo1.ListIndex = 1 ' Select SCAN_CURRENT_USER on load

    Combo2.Text = "Software\Microsoft"
    Combo3.Text = 2 ' Scan through all subdirectories
    
    Combo4.AddItem "SCAN_ADDREMOVELISTS", 0
    Combo4.AddItem "SCAN_CUSTOMCONTROLS", 1
    Combo4.AddItem "SCAN_FILEEXTENSIONS", 2
    Combo4.AddItem "SCAN_HELPRESOURCES", 3
    Combo4.AddItem "SCAN_SHAREDDLLS", 4
    Combo4.AddItem "SCAN_SHELLFOLDERS", 5
    Combo4.AddItem "SCAN_SOFTWAREPATHS", 6
    Combo4.AddItem "SCAN_STARTUPKEYS", 7
    Combo4.AddItem "SCAN_WINDOWSFONTS", 8
    
    Combo4.ListIndex = 1
    
    Combo5.Text = ""
    Combo6.Text = ""
    
    Check1.Value = vbChecked
    Check2_Click
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Begin Scanning" Then
        ListView1.ListItems.Clear
        t = Timer
    
        Select Case Combo1.ListIndex
            Case 0: SCANREG.ClassRoot = SCAN_CLASSES_ROOT
            Case 1: SCANREG.ClassRoot = SCAN_CURRENT_USER
            Case 2: SCANREG.ClassRoot = SCAN_LOCAL_MACHINE
        End Select
        
        SCANREG.FilterData = Combo5.Text
        SCANREG.FilterKeys = Combo6.Text
        SCANREG.FullRegistryScan = (Check8.Value = vbChecked)
        SCANREG.ScanPath = Combo2.Text
        SCANREG.ScanSubKeys = (Check1.Value = vbChecked)
        SCANREG.ScanDeep = Combo3.Text
        SCANREG.CustomScan = (Check2.Value = vbChecked)
        SCANREG.CustomScanPath = Combo4.ListIndex
        
        Command1.Caption = "Pause Scanning"
        Command2.Enabled = True
        SCANREG.BeginScanning
    ElseIf Command1.Caption = "Pause Scanning" Then
        Command1.Caption = "Resume Scanning"
        SCANREG.PauseScanning
    Else
        Command1.Caption = "Pause Scanning"
        SCANREG.ResumeScanning
    End If
End Sub

Private Sub Command2_Click()
    SCANREG.CancelScanning
End Sub

Private Sub Check1_Click()
    If Check8.Value = vbChecked Then
        Combo3.Locked = True
        Combo3.BackColor = vbButtonFace
    Else
        Combo3.Locked = (Check1.Value <> vbChecked)
        If Combo3.Locked Then
            Combo3.BackColor = vbButtonFace
        Else
            Combo3.BackColor = vbWhite
        End If
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        Combo1.Locked = True
        Combo1.BackColor = vbButtonFace
        Combo2.Locked = True
        Combo2.BackColor = vbButtonFace
        Combo4.Locked = False
        Combo4.BackColor = vbWindowBackground
        Check1.Enabled = False
    Else
        Combo1.Locked = False
        Combo1.BackColor = vbWindowBackground
        Combo2.Locked = False
        Combo2.BackColor = vbWindowBackground
        Combo4.Locked = True
        Combo4.BackColor = vbButtonFace
        Check1.Enabled = True
    End If
    Check1_Click
End Sub

Private Sub Check3_Click()
    Check4.Enabled = (Check3.Value = vbChecked)
    If Check3.Value = vbChecked Then
        Check5.Enabled = (Check4.Value = vbChecked)
        Check6.Enabled = (Check4.Value = vbChecked)
    Else
        Check5.Enabled = False
        Check6.Enabled = False
    End If
End Sub

Private Sub Check4_Click()
    Check5.Enabled = (Check4.Value = vbChecked)
    Check6.Enabled = (Check4.Value = vbChecked)
End Sub

Private Sub Check8_Click()
    Check2.Enabled = (Check8.Value <> vbChecked)
    If Check8.Value = vbChecked Then
        Check2.Value = vbChecked
        Check2_Click
        Combo4.Locked = False
        Combo4.BackColor = vbButtonFace
        Command1.SetFocus
    Else
        Check2.Value = vbUnchecked
        Check2_Click
        Combo2.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        SCANREG.CancelScanning
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SCANREG = Nothing
    Set REG = Nothing
End Sub
