VERSION 5.00
Begin VB.Form frmJournalRecordHook 
   Caption         =   "Journal Record Hook Demonstration"
   ClientHeight    =   6015
   ClientLeft      =   5730
   ClientTop       =   3015
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJournalRecordHook.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   6585
   Begin VB.PictureBox picInfo 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   6375
      TabIndex        =   2
      Top             =   4860
      Width           =   6375
      Begin VB.Label lblIdle 
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   1380
         TabIndex        =   8
         Top             =   0
         Width           =   4995
      End
      Begin VB.Label lblPixelsMoved 
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   1380
         TabIndex        =   7
         Top             =   360
         Width           =   4995
      End
      Begin VB.Label lblKeysPressed 
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   1380
         TabIndex        =   6
         Top             =   720
         Width           =   4995
      End
      Begin VB.Label lblInfo 
         Caption         =   "UI Idle for:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Caption         =   "Pixels travelled:"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Caption         =   "Keys pressed:"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   720
         Width           =   1275
      End
   End
   Begin VB.Timer tmrIdleUI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   0
   End
   Begin VB.ListBox lstEvents 
      Height          =   4350
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6375
   End
   Begin VB.CheckBox chkHook 
      Caption         =   "&Hook"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmJournalRecordHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_SYSTEMKEYDOWN = &H104
Private Const WM_SYSTEMKEYUP = &H105

Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private m_lLastTime As Long
Private m_tP As POINTAPI
Private m_lPixels As Long
Private m_lKeys As Long

Implements IWindowsHook

Private Sub updateIdleUI()
   lblIdle.Caption = timeGetTime() - m_lLastTime & "ms"
   lblPixelsMoved.Caption = m_lPixels
   lblKeysPressed.Caption = m_lKeys
End Sub

Private Sub enableUIControls(ByVal bState As Boolean)
Dim lColor As OLE_COLOR
Dim i As Long
   lColor = IIf(bState, vbWindowText, vbGrayText)
   lblIdle.ForeColor = lColor
   lblPixelsMoved.ForeColor = lColor
   lblKeysPressed.ForeColor = lColor
   For i = lblInfo.LBound To lblInfo.UBound
      lblInfo(i).ForeColor = lColor
   Next i
End Sub

Private Function ClassName(ByVal hWnd As Long) As String
Dim sBuf As String
Dim iPos As Long
   sBuf = String$(256, 0)
   GetClassName hWnd, sBuf, 255
   iPos = InStr(sBuf, vbNullChar)
   If (iPos > 0) Then
      ClassName = Left(sBuf, iPos - 1)
   Else
      ClassName = sBuf
   End If
End Function

Private Function shiftState() As String
Dim sRet As String
   sRet = "  "
   If Not (GetAsyncKeyState(vbKeyControl) = 0) Then
      sRet = sRet & "Ctrl"
   End If
   If Not (GetAsyncKeyState(vbKeyMenu) = 0) Then
      If Len(sRet) > 0 Then
         sRet = sRet & "+"
      End If
      sRet = sRet & "Alt"
   End If
   If Not (GetAsyncKeyState(vbKeyShift) = 0) Then
      If Len(sRet) > 0 Then
         sRet = sRet & "+"
      End If
      sRet = sRet & "Shift"
   End If
   shiftState = sRet
End Function

Private Sub LogEvent(ByVal sMsg As String)
   m_lLastTime = timeGetTime()
   updateIdleUI
   lstEvents.AddItem sMsg
   lstEvents.ListIndex = lstEvents.ListCount - 1
End Sub

Private Sub chkHook_Click()
   If (chkHook.Value = Checked) Then
      InstallHook Me, WH_JOURNALRECORD
      InstallHook Me, WH_GETMESSAGE
      timeBeginPeriod 1
      m_lLastTime = timeGetTime()
      tmrIdleUI.Enabled = True
      enableUIControls True
   Else
      enableUIControls False
      tmrIdleUI.Enabled = False
      timeEndPeriod 1
      RemoveHook Me, WH_GETMESSAGE
      RemoveHook Me, WH_JOURNALRECORD
   End If
End Sub

Private Sub Form_Load()
   enableUIControls False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If (chkHook.Value = Checked) Then
      chkHook.Value = Unchecked
   End If
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   picInfo.Move picInfo.Left, Me.ScaleHeight - picInfo.Height - picInfo.Left, _
      Me.ScaleWidth - picInfo.Left * 2
   lstEvents.Move picInfo.Left, lstEvents.Top, _
      Me.ScaleWidth - picInfo.Left * 2, Me.ScaleHeight - lstEvents.Top - picInfo.Height - picInfo.Left
End Sub

Private Function IWindowsHook_HookProc( _
      ByVal eType As vbalWinHook6.EHTHookTypeConstants, _
      ByVal nCode As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long, _
      bConsume As Boolean _
   ) As Long

   If (eType = WH_GETMESSAGE) Then
      Dim cMsg As cGetMsglParam
      On Error Resume Next
      Set cMsg = GetMsglParam(lParam)
      If (Err.Number) Then
         Debug.Print Err.Description
      End If
      
      If (cMsg.Message = WM_CANCELJOURNAL) Then
         LogEvent "Journalling cancelled by Ctrl+Esc or Ctrl+Alt+Del; Removing Hook"
         If (chkHook.Value = Checked) Then
            chkHook.Value = Unchecked
         End If
      End If
      
   
   ElseIf (eType = WH_JOURNALRECORD) Then
   
      Dim keyCode As Integer
      Dim keyAscii As Integer
      Dim sMsg As String
      Dim sMsgBit As String
      Dim lhWnd As Long
      Dim tP As POINTAPI
      
      Dim cEvent As cJournallParam
      Set cEvent = JournalRecordlParam(lParam)
      
      lhWnd = cEvent.hWnd
      If Not (lhWnd = 0) Then
         sMsgBit = "  over Window " & _
            Hex$(lhWnd) & _
            " (" & ClassName(lhWnd) & ")"
      End If
            
      Select Case cEvent.Msg
      
      Case WM_KEYDOWN
         m_lLastTime = timeGetTime()
         m_lKeys = m_lKeys + 1
         updateIdleUI
         
         keyCode = (cEvent.lParamLow And &HFF&)
         LogEvent "KeyDown " & keyCode & shiftState & sMsgBit
         
      Case WM_SYSTEMKEYDOWN
         m_lLastTime = timeGetTime()
         m_lKeys = m_lKeys + 1
         updateIdleUI
      
         keyCode = (cEvent.lParamLow And &HFF&)
         If Not (keyCode = vbKeyMenu) Then
            LogEvent "SystemKeyDown " & keyCode & shiftState & sMsgBit
         End If
      
      Case WM_KEYUP
         m_lLastTime = timeGetTime()
         updateIdleUI
      
         keyCode = (cEvent.lParamLow And &HFF&)
         If Not (keyCode = vbKeyMenu) Then
            LogEvent "KeyUp " & keyCode & shiftState & sMsgBit
         End If
      
      Case WM_SYSTEMKEYUP
         m_lLastTime = timeGetTime()
         updateIdleUI
         
         keyCode = (cEvent.lParamLow And &HFF&)
         LogEvent "SystemKeyUp " & keyCode & shiftState & sMsgBit
      
      
      Case WM_MOUSEMOVE
         GetCursorPos tP
         If Not (((m_tP.x - tP.x) = 0) And ((m_tP.y - tP.y) = 0)) Then
            m_lPixels = m_lPixels + Sqr((m_tP.x - tP.x) ^ 2 + (m_tP.y - tP.y) ^ 2)
            m_lLastTime = timeGetTime()
            updateIdleUI
            LSet m_tP = tP
         End If
               
      Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
         m_lLastTime = timeGetTime()
         updateIdleUI
         
         sMsg = "Mouse Down: "
         Select Case cEvent.Msg
         Case WM_LBUTTONDOWN
            sMsg = sMsg & "Left Button"
         Case WM_RBUTTONDOWN
            sMsg = sMsg & "Right Button"
         Case WM_MBUTTONDOWN
            sMsg = sMsg & "Middle Button"
         End Select
         LogEvent sMsg & ", x=" & cEvent.lParamLow & ",y=" & cEvent.lParamHigh & shiftState & sMsgBit
         
      
      Case WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP
         m_lLastTime = timeGetTime()
         updateIdleUI
         
         sMsg = "Mouse Up: "
         Select Case cEvent.Msg
         Case WM_LBUTTONUP
            sMsg = sMsg & "Left Button"
         Case WM_RBUTTONUP
            sMsg = sMsg & "Right Button"
         Case WM_MBUTTONUP
            sMsg = sMsg & "Middle Button"
         End Select
         LogEvent sMsg & ", x=" & cEvent.lParamLow & ",y=" & cEvent.lParamHigh & shiftState & sMsgBit
         
      Case WM_MOUSEWHEEL
         m_lLastTime = timeGetTime()
         updateIdleUI
                  
         ' it does not seem to be possible to determine what mouse wheel action was being taken
         LogEvent "MouseWheel" & ", x=" & cEvent.lParamLow & ",y=" & cEvent.lParamHigh & shiftState & sMsgBit
      
      Case WM_LBUTTONDBLCLK, WM_RBUTTONDBLCLK, WM_MBUTTONDBLCLK
         m_lLastTime = timeGetTime()
         updateIdleUI
         
         sMsg = "Double Click: "
         Select Case cEvent.Msg
         Case WM_LBUTTONDBLCLK
            sMsg = sMsg & "Left Button"
         Case WM_RBUTTONDBLCLK
            sMsg = sMsg & "Right Button"
         Case WM_MBUTTONDBLCLK
            sMsg = sMsg & "Middle Button"
         End Select
         LogEvent sMsg & ", x=" & cEvent.lParamLow & ",y=" & cEvent.lParamHigh & shiftState & sMsgBit
               
      End Select
               
   End If
End Function

Private Sub tmrIdleUI_Timer()
   updateIdleUI
End Sub
