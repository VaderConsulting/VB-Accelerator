VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "vbAccelerator Goldfish"
   ClientHeight    =   4230
   ClientLeft      =   5415
   ClientTop       =   4440
   ClientWidth     =   5055
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2919.621
   ScaleMode       =   0  'User
   ScaleWidth      =   4746.906
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3540
      TabIndex        =   7
      Top             =   3780
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame fraSep 
      Height          =   75
      Left            =   -60
      TabIndex        =   6
      Top             =   3180
      Width           =   5835
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   720
      Left            =   60
      ScaleHeight     =   463.54
      ScaleMode       =   0  'User
      ScaleWidth      =   3381.735
      TabIndex        =   1
      Top             =   60
      Width           =   4875
      Begin VB.Image imgVB 
         Height          =   660
         Left            =   2520
         Picture         =   "fAbout.frx":0000
         Top             =   0
         Width           =   660
      End
      Begin VB.Image Image1 
         Height          =   660
         Left            =   0
         Picture         =   "fAbout.frx":0488
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
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
      Height          =   405
      Left            =   3540
      TabIndex        =   0
      Top             =   3300
      Width           =   1380
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   345
      Left            =   2760
      TabIndex        =   5
      Top             =   1020
      Width           =   2145
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"fAbout.frx":0D11
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   810
      Index           =   0
      Left            =   1020
      TabIndex        =   2
      Top             =   1560
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"fAbout.frx":0DD8
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   60
      TabIndex        =   3
      Top             =   3345
      Width           =   3270
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = "About " & App.Title
   Me.Icon = frmIconEx.Icon
   lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   lblTitle.Caption = App.Title
   ' centre to owner:
   Dim lLeft As Long, lTop As Long
   With frmIconEx
     lLeft = .left + (.Width - Me.Width) \ 2
     lTop = .top + (.Height - Me.Height) \ 2
   End With
   If (lLeft + Me.Width) > Screen.Width Then lLeft = Screen.Width - Me.Width
   If (lTop + Me.Height) > Screen.Height Then lTop = Screen.Height - Me.Height
   If (lLeft < 0) Then lLeft = 0
   If (lTop < 0) Then lTop = 0
   Me.Move lLeft, lTop
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    'Dim cR As New cRegistry
    'cR.ClassKey = HKEY_LOCAL_MACHINE
    'cR.SectionKey = gREGKEYSYSINFO
    'cR.ValueKey = gREGVALSYSINFO
    'cR.Default = "!Not Found!"
    'If (cR.Value <> "!Not Found!") Then
    '    SysInfoPath = cR.Value
    'Else
    '    cR.SectionKey = gREGKEYSYSINFOLOC
    '    cR.ValueKey = gREGVALSYSINFOLOC
    '    If (cR.Value <> "!Not Found!") Then
    '        SysInfoPath = cR.Value
    '        ' Validate Existance Of Known 32 Bit File Version
    '        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
    '            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
    '        ' Error - File Can Not Be Found...
    '        Else
    '            GoTo SysInfoErr
    '        End If
    '    Else
    '        GoTo SysInfoErr
    '    End If
    'End If
    '
    'Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

