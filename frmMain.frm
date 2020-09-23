VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmGlobal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Rename: Copyright (c) Neil Etherington 2000"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   28
      Top             =   8520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CheckBox chkSCase 
      Caption         =   "Selected Case:"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   27
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CheckBox chkLCase 
      Caption         =   "Lower Case:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7560
      TabIndex        =   26
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CheckBox chkUCase 
      Caption         =   "Upper Case:"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   25
      Top             =   6120
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox txtSpec 
      Height          =   285
      Left            =   7560
      TabIndex        =   5
      Text            =   "*.*"
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CheckBox chkDecrease 
      Caption         =   "Decrease Value:"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CheckBox chkIncrease 
      Caption         =   "Increase Value:"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin ComctlLib.ProgressBar pbrFiles 
      Height          =   135
      Left            =   240
      TabIndex        =   18
      Top             =   945
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txtNamed 
      Height          =   285
      Left            =   8880
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "Select Operation:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8400
      Width           =   2415
   End
   Begin VB.TextBox txtRenameName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   0
      Text            =   "X-File"
      ToolTipText     =   "Rename Files To:"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtRnameCount 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Start Count From This Number:"
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox txtReExt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   4
      Text            =   ".txt"
      ToolTipText     =   "Extension For Renamed Files:"
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox txtReFileCount 
      Height          =   285
      Left            =   7560
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "0"
      ToolTipText     =   "#Count:"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Hidden          =   -1  'True
      Left            =   360
      System          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   8520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   6615
      Left            =   360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "File Selection:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   7095
      Begin VB.ListBox SBList1 
         Height          =   7080
         Left            =   3240
         MultiSelect     =   2  'Extended
         TabIndex        =   29
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Progress:"
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   8895
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rename:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   630
      Index           =   1
      Left            =   3720
      TabIndex        =   24
      Top             =   0
      Width           =   1950
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Rename:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   0
      Left            =   3690
      TabIndex        =   23
      Top             =   0
      Width           =   1950
   End
   Begin VB.Label lblOp 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operation:"
      Height          =   255
      Left            =   7320
      TabIndex        =   22
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label lblSpec 
      Caption         =   "FileSpec:"
      Height          =   255
      Left            =   7560
      TabIndex        =   21
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label lblPB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9120
      TabIndex        =   20
      Top             =   870
      Width           =   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Files Named:"
      Height          =   195
      Left            =   8880
      TabIndex        =   17
      Top             =   5400
      Width           =   915
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      Index           =   2
      X1              =   7335
      X2              =   7335
      Y1              =   1815
      Y2              =   9000
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   10215
      X2              =   10215
      Y1              =   1845
      Y2              =   9000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   2
      X1              =   7320
      X2              =   10200
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   7320
      X2              =   10200
      Y1              =   1845
      Y2              =   1845
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   10200
      X2              =   10200
      Y1              =   1830
      Y2              =   9000
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   7320
      X2              =   10200
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Line Line2 
      X1              =   7320
      X2              =   10200
      Y1              =   8990
      Y2              =   8990
   End
   Begin VB.Line Line1 
      X1              =   7320
      X2              =   7320
      Y1              =   1845
      Y2              =   9000
   End
   Begin VB.Label lblRenameName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   7560
      TabIndex        =   15
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label lblRnameCounter 
      AutoSize        =   -1  'True
      Caption         =   "Counter: (Start From This Value)"
      Height          =   195
      Left            =   7560
      TabIndex        =   14
      Top             =   3960
      Width           =   2250
   End
   Begin VB.Label lblReExten 
      AutoSize        =   -1  'True
      Caption         =   "Extension:"
      Height          =   195
      Left            =   7560
      TabIndex        =   13
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblReFileCount 
      AutoSize        =   -1  'True
      Caption         =   "File Count:"
      Height          =   195
      Left            =   7560
      TabIndex        =   12
      Top             =   5400
      Width           =   750
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuh1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuh2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileSelectNil 
         Caption         =   "Select Nil"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuh3 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRename 
      Caption         =   "&Rename"
      Begin VB.Menu mnuh4 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuRRenamefile 
         Caption         =   "Rename File"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuh6 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuRRenameExtension 
         Caption         =   "Rename Extension"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'form load
Private Sub Form_Load()
    
    'list files
    subListFiles

End Sub

'form unload
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Close
End Sub

'drive1 change
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

'dir1 change
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    subListFiles
End Sub

'dir1 path
Private Sub Dir1_Click()
    Dim x As Integer
    x = Dir1.ListIndex
    Dir1.Path = Dir1.List(x)
End Sub

'determines which checkbox is enabled
Private Sub chkDecrease_Click()
    If chkDecrease.Value = vbChecked Then
        chkIncrease.Value = vbUnchecked
    End If
End Sub

'determines which checkbox is enabled
Private Sub chkIncrease_Click()
    If chkIncrease.Value = vbChecked Then
        chkDecrease.Value = vbUnchecked
    End If
End Sub

'rename files
Private Sub cmdRename_Click()
        
    If lblOp.Caption = "Rename File" Then
        modRename.subRename
    ElseIf lblOp.Caption = "Rename Extension" Then
        modRenameExt.subReanmeExt
    End If
    
End Sub

Private Sub Form_Resize()

On Error Resume Next

With File1
    .Height = frmGlobal.Height - 6000
End With

End Sub

'list files
Public Sub subListFiles()

    On Error Resume Next

    Dim x As Integer

    SBList1.Clear
    File1.Refresh
    File1.ListIndex = 0
    SBList1.ListIndex = 0

    For x = 0 To File1.ListCount - 1
        
        File1.ListIndex = (x)
        SBList1.ListIndex = (x)
        
        SBList1.AddItem File1.FileName
    Next x


End Sub

'close application
Private Sub mnuFileExit_Click()
    Unload Me
    Close
    End
End Sub

'select all
Private Sub mnuFileSelectAll_Click()
    Dim x As Integer
    
    SBList1.ListIndex = 0
    
    For x = 0 To SBList1.ListCount - 1
        SBList1.ListIndex = (x)
                
        SBList1.Selected(SBList1.ListIndex) = True
    Next x
    
End Sub

'select nil
Private Sub mnuFileSelectNil_Click()
    Dim x As Integer
    
    SBList1.ListIndex = 0
    
    For x = 0 To SBList1.ListCount - 1
        SBList1.ListIndex = (x)
                
        SBList1.Selected(SBList1.ListIndex) = False
    Next x
    
End Sub

Private Sub SBList1_Click()
    File1.ListIndex = SBList1.ListIndex
End Sub

Private Sub txtReExt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdRename_Click
    End If
End Sub

'rename files on keydown on name
Private Sub txtRenameName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdRename_Click
    End If
End Sub

'filespec
Private Sub txtSpec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        File1.Pattern = txtSpec.Text
        subListFiles
    End If
End Sub

'rename file
Private Sub mnuRRenamefile_Click()

    lblOp.Caption = "Rename File"
    cmdRename.Caption = "Rename File"
    chkSCase.Enabled = False
    chkUCase.Enabled = False
    chkLCase.Enabled = False
    txtRenameName.Enabled = True
    chkIncrease.Enabled = True
    chkDecrease.Enabled = True
    txtRnameCount.Enabled = True
    txtReExt.Enabled = True
    txtRenameName.SetFocus
    
End Sub

'rename extension
Private Sub mnuRRenameExtension_Click()

    lblOp.Caption = "Rename Extension"
    cmdRename.Caption = "Rename Extension"
    chkSCase.Enabled = True
    chkUCase.Enabled = True
    chkLCase.Enabled = True
    txtRenameName.Enabled = False
    chkIncrease.Enabled = False
    chkDecrease.Enabled = False
    txtRnameCount.Enabled = False
    txtReExt.Enabled = True
    txtReExt.SetFocus

End Sub

'lcase
Private Sub chkLCase_Click()
    chkUCase.Value = vbUnchecked
    chkSCase.Value = vbUnchecked
End Sub

'ucase
Private Sub chkUCase_Click()
    chkLCase.Value = vbUnchecked
    chkSCase.Value = vbUnchecked
End Sub

'scase
Private Sub chkSCase_Click()
    chkLCase.Value = vbUnchecked
    chkUCase.Value = vbUnchecked
End Sub

