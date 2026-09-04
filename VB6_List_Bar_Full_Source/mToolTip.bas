Attribute VB_Name = "mToolTip"
Option Explicit

Public Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uId As Long
    rct As RECT
    hinst As Long
    lpszText As Long
End Type

Public Type ToolTipText
    hdr As NMHDR
    lpszText As Long
    szText As String * 80
    hinst As Long
    uFlags As Long
End Type

Private Const H_MAX As Long = &HFFFF + 1

Private Const WM_USER = &H400&
Public Const TTM_RELAYEVENT = (WM_USER + 7)
Private Const TTM_ACTIVATE = (WM_USER + 1)
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ADDTOOL = TTM_ADDTOOLA
Private Const TTM_DELTOOLA = (WM_USER + 5)
Private Const TTM_DELTOOL = TTM_DELTOOLA

Private Const TTN_FIRST = (H_MAX - 520&)
Public Const TTN_NEEDTEXTA = (TTN_FIRST - 0&)
Public Const TTN_NEEDTEXT = TTN_NEEDTEXTA

Private Const TOOLTIPS_CLASS = "tooltips_class32"
Private Const TTF_IDISHWND = &H1
Private Const LPSTR_TEXTCALLBACK As Long = -1

Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()

' Tooltips:
Private m_hWndToolTip As Long
Private m_iRef As Long
Public gsToolTipBuffer As String         'Tool tip text; This string must have
                                         'module or global level scope, because
                                         'a pointer to it is copied into a
                                         'ToolTipText structure
Public gsInfoTipBuffer As String

Public Property Get hwndToolTip() As Long
   If m_hWndToolTip = 0 Then
      CreateToolTip
   End If
   hwndToolTip = m_hWndToolTip
End Property
Public Sub AddToToolTip(ByVal hWnd As Long)
Dim tTi As TOOLINFO

   If m_hWndToolTip = 0 Then
      CreateToolTip
      If m_hWndToolTip = 0 Then
         Exit Sub
      End If
   End If
    
   With tTi
      .cbSize = Len(tTi)
      .uId = hWnd
      .hWnd = hWnd
      .hinst = App.hInstance
      .uFlags = TTF_IDISHWND
      .lpszText = LPSTR_TEXTCALLBACK
   End With
   
   SendMessage m_hWndToolTip, TTM_ADDTOOL, 0, tTi
   SendMessageLong m_hWndToolTip, TTM_ACTIVATE, 1, hWnd
   m_iRef = m_iRef + 1

End Sub
Public Sub RemoveFromToolTip(ByVal hWnd As Long)
Dim tTi As TOOLINFO
   If m_hWndToolTip <> 0 Then
      With tTi
         .cbSize = Len(tTi)
         .uId = hWnd
         .hWnd = hWnd
      End With
      SendMessage m_hWndToolTip, TTM_DELTOOL, 0, tTi
      
      m_iRef = m_iRef - 1
      If m_iRef <= 0 Then
         DestroyWindow m_hWndToolTip
         m_hWndToolTip = 0
         m_iRef = 0
      End If
   End If
End Sub
 
Private Sub CreateToolTip()
   ' Create the tooltip:
   InitCommonControls
   m_hWndToolTip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, vbNullString, 0, _
             CW_USEDEFAULT, CW_USEDEFAULT, _
             CW_USEDEFAULT, CW_USEDEFAULT, _
             0, 0, _
             App.hInstance, _
             ByVal 0)
   SendMessage m_hWndToolTip, TTM_ACTIVATE, 1, ByVal 0
End Sub

