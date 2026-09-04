VERSION 5.00
Begin VB.Form frmMonitors 
   Caption         =   "System Monitors"
   ClientHeight    =   4410
   ClientLeft      =   2805
   ClientTop       =   2295
   ClientWidth     =   7215
   Icon            =   "frmTestMonitors.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7215
   Begin VB.PictureBox picTools 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   7215
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmdMaximise 
         Caption         =   "Ma&ximise"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdCentre 
         Caption         =   "&Centre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2940
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "&Move"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.ComboBox cboMonitors 
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
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   30
         Width           =   1815
      End
   End
   Begin VB.Timer tmrMove 
      Interval        =   50
      Left            =   6660
      Top             =   720
   End
   Begin VB.PictureBox picMonitors 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   60
      ScaleHeight     =   159.75
      ScaleMode       =   0  'User
      ScaleWidth      =   321.75
      TabIndex        =   0
      Top             =   480
      Width           =   6495
   End
End
Attribute VB_Name = "frmMonitors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cM As New cMonitors

Private Sub showMonitors()
Dim i As Long

   With picMonitors
      .Cls
      .ScaleWidth = m_cM.VirtualScreenWidth + 64
      .ScaleHeight = .Height * (.ScaleWidth / .Width)
           
   End With
   
   With m_cM
      picMonitors.Line (.VirtualScreenLeft + 32, .VirtualScreenTop + 32)-(.VirtualScreenWidth + .VirtualScreenLeft + 32, .VirtualScreenHeight + .VirtualScreenTop + 32), &HFF00&, BF
      For i = 1 To .MonitorCount
         With .Monitor(i)
            picMonitors.ForeColor = vbWhite
            picMonitors.Line (.Left + 32, .Top + 32)-(.Left + 32 + .Width, .Top + 32 + .Height), QBColor(i * 2), BF
            picMonitors.Line (.WorkLeft + 32, .WorkTop + 32)-(.WorkLeft + 32 + .WorkWidth, .WorkTop + 32 + .WorkHeight), QBColor(i * 2 + 1), BF
            
            picMonitors.CurrentX = .Left + 36
            picMonitors.CurrentY = .Top + 36
            picMonitors.Print .Name & IIf(.IsPrimary, " (Primary)", "")
            picMonitors.CurrentX = .Left + 36
         End With
      Next i
      
   End With

Dim lX1 As Long, lY1 As Long, lX2 As Long, lY2 As Long
   lX1 = Me.ScaleX(Me.Left, vbTwips, vbPixels) + 32
   lY1 = Me.ScaleY(Me.Top, vbTwips, vbPixels) + 32
   lX2 = lX1 + Me.ScaleX(Me.Width, vbTwips, vbPixels)
   lY2 = lY1 + Me.ScaleY(Me.Height, vbTwips, vbPixels)
   picMonitors.Line (lX1, lY1)-(lX2, lY2), vbWhite, BF
   
   picMonitors.CurrentX = 32
   picMonitors.CurrentY = m_cM.VirtualScreenHeight + m_cM.VirtualScreenHeight \ 32
   picMonitors.Print IIf(m_cM.AllMonitorsSame, "All monitors have same display format", "Monitors have different display formats")
   
   picMonitors.Refresh
   
End Sub

Private Sub cmdCentre_Click()
Dim cMonTo As cMonitor
Dim lWidth As Long
Dim lHeight As Long
   
   Set cMonTo = m_cM.Monitor(cboMonitors.ItemData(cboMonitors.ListIndex))
   
   If (Me.WindowState = vbMaximized) Then
      Me.WindowState = vbNormal
   End If
   
   lWidth = Me.ScaleX(Me.Width, vbTwips, vbPixels)
   If (lWidth > cMonTo.WorkWidth) Then
      lWidth = cMonTo.WorkWidth
   End If
   lHeight = Me.ScaleY(Me.Height, vbTwips, vbPixels)
   If (lHeight > cMonTo.WorkHeight) Then
      lHeight = cMonTo.WorkHeight
   End If
   
   Me.Move _
      Me.ScaleX((cMonTo.WorkLeft + (cMonTo.WorkWidth - lWidth) \ 2), vbPixels, vbTwips), _
      Me.ScaleY((cMonTo.WorkTop + (cMonTo.WorkHeight - lHeight) \ 2), vbPixels, vbTwips)
   
   '
End Sub

Private Sub cmdMaximise_Click()
   '
Dim cMonTo As cMonitor
Dim cMonFrom As cMonitor
   
   Set cMonTo = m_cM.Monitor(cboMonitors.ItemData(cboMonitors.ListIndex))
   Set cMonFrom = m_cM.MonitorForWindow(Me.hwnd)
   
   If Not (cMonTo.hMonitor = cMonFrom.hMonitor) Then
      ' we need to swap onto the new monitor:
      If (Me.WindowState = vbMaximized) Then
         Me.WindowState = vbNormal
      End If
      Me.Move Me.ScaleX(cMonTo.Left, vbPixels, vbTwips), Me.ScaleY(cMonTo.Top, vbPixels, vbTwips)
   End If
   Me.WindowState = vbMaximized

   '
End Sub

Private Sub cmdMove_Click()
   '
Dim cMonTo As cMonitor
Dim cMonFrom As cMonitor
   
   Set cMonTo = m_cM.Monitor(cboMonitors.ItemData(cboMonitors.ListIndex))
   Set cMonFrom = m_cM.MonitorForWindow(Me.hwnd)
   
   If Not (cMonTo.hMonitor = cMonFrom.hMonitor) Then
      
      ' work out my offset from the origin of this monitor:
      Dim lOffsetX As Long
      Dim lOffsetY As Long
      lOffsetX = ScaleX(Me.Left, vbTwips, vbPixels) - cMonFrom.Left
      lOffsetY = ScaleY(Me.Top, vbTwips, vbPixels) - cMonFrom.Top
      
      If (lOffsetX < 0) Then
         lOffsetX = 32
      End If
      If (lOffsetY < 0) Then
         lOffsetY = 32
      End If
      
      ' move to the newly selected monitor
      Me.Move Me.ScaleX(lOffsetX + cMonTo.Left, vbPixels, vbTwips), Me.ScaleY(lOffsetY + cMonTo.Top, vbPixels, vbTwips)
         
   End If
   '
End Sub

Private Sub Form_Load()
Dim i As Long
   With m_cM
      For i = 1 To .MonitorCount
         cboMonitors.AddItem .Monitor(i).Name
         cboMonitors.ItemData(cboMonitors.NewIndex) = i
      Next i
   End With
   If (cboMonitors.ListCount > 0) Then
      cboMonitors.ListIndex = 0
   End If
End Sub

Private Sub Form_Paint()
   showMonitors
End Sub

Private Sub Form_Resize()
On Error Resume Next
   picMonitors.Move picMonitors.Left, picMonitors.Top, _
      Me.ScaleWidth - picMonitors.Left * 2, _
      Me.ScaleHeight - picMonitors.Top - picMonitors.Left
End Sub

Private Sub tmrMove_Timer()
Static fLeft As Single, fTop As Single
   ' really use WM_MOVING for this, but for the sake of the demo:
   If Not (Me.Top = fLeft) Or Not (Me.Left = fLeft) Then
      Me.Refresh
   End If
End Sub
