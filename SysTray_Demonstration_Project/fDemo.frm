VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SysTray Sample Application"
   ClientHeight    =   4125
   ClientLeft      =   4920
   ClientTop       =   1785
   ClientWidth     =   5790
   Icon            =   "fDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdShowBalloon 
      Caption         =   "Show In&foTip"
      Height          =   495
      Left            =   180
      TabIndex        =   6
      Top             =   3420
      Width           =   1395
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "&Stimpy"
      Height          =   375
      Index           =   2
      Left            =   660
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "&Tooth Beaver"
      Height          =   375
      Index           =   1
      Left            =   660
      TabIndex        =   4
      Top             =   2220
      Width           =   1335
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "&Bob"
      Height          =   375
      Index           =   0
      Left            =   660
      TabIndex        =   3
      Top             =   1620
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CheckBox chkSysTray 
      Caption         =   "&Show in Systray"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   120
      Picture         =   "fDemo.frx":08CA
      Top             =   2700
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "fDemo.frx":148C
      Top             =   2100
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "fDemo.frx":1D56
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image imgLogo 
      Height          =   660
      Left            =   120
      Picture         =   "fDemo.frx":2620
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblDetail 
      BackStyle       =   0  'Transparent
      Caption         =   "VB Source Code and Tips at http://vbaccelerator.com/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   2835
   End
   Begin VB.Label lblBlack 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   5655
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1

Private Sub SetIcon()
    Select Case True
    Case optIcon(0).Value
        m_frmSysTray.IconHandle = imgIcon(0).Picture.Handle
    Case optIcon(1).Value
        m_frmSysTray.IconHandle = imgIcon(1).Picture.Handle
    Case optIcon(2).Value
        m_frmSysTray.IconHandle = imgIcon(2).Picture.Handle
    End Select
End Sub

Private Sub chkSysTray_Click()
    If (chkSysTray.Value = Checked) Then
        Set m_frmSysTray = New frmSysTray
        With m_frmSysTray
            .AddMenuItem "&Open SysTray Sample", "open", True
            .AddMenuItem "-"
            .AddMenuItem "&vbAccelerator on the Web", "vbAccelerator"
            .AddMenuItem "&About...", "About"
            .AddMenuItem "-"
            .AddMenuItem "&Close", "close"
            .ToolTip = "SysTray Sample!"
        End With
        SetIcon
    Else
        Unload m_frmSysTray
        Set m_frmSysTray = Nothing
    End If
End Sub

Private Sub cmdShowBalloon_Click()
   m_frmSysTray.ShowBalloonTip _
      "Hello from vbAccelerator.com.  This SysTray form allows Unicode text and balloon tips.", _
      "vbAccelerator SysTray Sample", _
      NIIF_INFO
End Sub

Private Sub Form_Load()
    chkSysTray_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload m_frmSysTray
    Set m_frmSysTray = Nothing
End Sub

Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
   Select Case sKey
   Case "open"
      Me.Show
      Me.ZOrder
   Case "close"
      Unload Me
   Case Else
      MsgBox "Clicked item with key " & sKey, vbInformation
   End Select
    
End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    Me.Show
    Me.ZOrder
End Sub

Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If (eButton = vbRightButton) Then
        m_frmSysTray.ShowMenu
    End If
End Sub

Private Sub optIcon_Click(Index As Integer)
    SetIcon
End Sub
