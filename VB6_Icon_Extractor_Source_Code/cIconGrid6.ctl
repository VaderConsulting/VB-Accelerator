VERSION 5.00
Begin VB.UserControl cIconGrid 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "cIconGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tIconInfo
   hIcon As Long
   lWidth As Long
   lHeight As Long
   lStartY As Long
   lColours As Double
   bSelected As Boolean
End Type
Private m_tI() As tIconInfo
Private m_iIconCount As Long
Private m_lWidth As Long
Private m_lHeight As Long
Private m_lTextHeight As Long
Private m_lStartX As Long
Private m_iSelRow As Long
Private m_bInFocus As Boolean
Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1

Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_TOP = &H0&
Private Const DT_LEFT = &H0&
Private Const DT_CENTER = &H1&
Private Const DT_RIGHT = &H2&
Private Const DT_VCENTER = &H4&
Private Const DT_BOTTOM = &H8&
Private Const DT_WORDBREAK = &H10&
Private Const DT_SINGLELINE = &H20&
Private Const DT_EXPANDTABS = &H40&
Private Const DT_TABSTOP = &H80&
Private Const DT_NOCLIP = &H100&
Private Const DT_EXTERNALLEADING = &H200&
Private Const DT_CALCRECT = &H400&
Private Const DT_NOPREFIX = &H800&
Private Const DT_INTERNAL = &H1000&
'#if(WINVER >= =&H0400)
Private Const DT_EDITCONTROL = &H2000&
Private Const DT_PATH_ELLIPSIS = &H4000&
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000

Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
' don't do it...
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get Selected(ByVal i As Long) As Boolean
   Selected = m_tI(i).bSelected
End Property

Public Property Let Selected(ByVal i As Long, ByVal bState As Boolean)
   m_tI(i).bSelected = bState
   m_iSelRow = i
   Draw
End Property

Public Property Get ItemCount() As Long
   ItemCount = m_iIconCount
End Property

Public Sub KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Dim i As Long
   Select Case KeyCode
   Case vbKeyUp
      If (m_iSelRow > 1) Then
         If (Shift And vbCtrlMask) = vbCtrlMask Then
            m_iSelRow = m_iSelRow - 1
         ElseIf (Shift And vbShiftMask) = vbShiftMask Then
            For i = m_iSelRow To m_iSelRow - 1 Step -1
               m_tI(i).bSelected = True
            Next i
            m_iSelRow = m_iSelRow - 1
         Else
            m_iSelRow = m_iSelRow - 1
            For i = 1 To m_iIconCount
               m_tI(i).bSelected = (i = m_iSelRow)
            Next i
         End If
         If Not (EnsureVisible(m_iSelRow)) Then
            Draw
         End If
      End If
   Case vbKeyDown
      If (m_iSelRow < m_iIconCount) Then
         If (Shift And vbCtrlMask) = vbCtrlMask Then
            m_iSelRow = m_iSelRow + 1
         ElseIf (Shift And vbShiftMask) = vbShiftMask Then
            For i = m_iSelRow To m_iSelRow + 1
               m_tI(i).bSelected = True
            Next i
            m_iSelRow = m_iSelRow + 1
         Else
            m_iSelRow = m_iSelRow + 1
            For i = 1 To m_iIconCount
               m_tI(i).bSelected = (i = m_iSelRow)
            Next i
         End If
         If Not (EnsureVisible(m_iSelRow)) Then
            Draw
         End If
      End If
      
   Case vbKeySpace
      If (m_iSelRow <> 0) Then
         m_tI(m_iSelRow).bSelected = Not (m_tI(m_iSelRow).bSelected)
         If Not (EnsureVisible(m_iSelRow)) Then
            Draw
         End If
      End If
   End Select
End Sub
Public Function EnsureVisible(ByVal iItem As Long) As Boolean
Dim lStartY As Long
Dim lOffset As Long
   If (m_cScroll.Visible(efsVertical)) Then
      lStartY = -m_cScroll.Value(efsVertical)
      If (m_tI(iItem).lStartY + m_tI(iItem).lHeight + m_lTextHeight + 6 + lStartY) > m_lHeight Then
         ' Can't see it, must scroll down:
         lOffset = (m_tI(iItem).lStartY + m_tI(iItem).lHeight + lStartY + m_lTextHeight + 6) - m_lHeight
         m_cScroll.Value(efsVertical) = m_cScroll.Value(efsVertical) + lOffset
      ElseIf (m_tI(iItem).lStartY + lStartY < 0) Then
         lOffset = (m_tI(iItem).lStartY + lStartY)
         m_cScroll.Value(efsVertical) = m_cScroll.Value(efsVertical) + lOffset
      End If
   End If
End Function
Public Sub MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim i As Long, j As Long, iItem As Long
Dim xL As Long, yL As Long
Dim lStartY As Long

   If (Button = 1) Then
      If (m_cScroll.Visible(efsVertical)) Then
         lStartY = -m_cScroll.Value(efsVertical)
      End If
      xL = X \ Screen.TwipsPerPixelX
      yL = Y \ Screen.TwipsPerPixelY
      ' Check which rect it is in:
      For i = 1 To m_iIconCount
         If (yL > m_tI(i).lStartY + lStartY) Then
            If (yL < m_tI(i).lStartY + lStartY + m_lTextHeight + 6 + m_tI(i).lHeight) Then
               iItem = i
               Exit For
            End If
         End If
      Next i
      
      If (iItem > 0) Then
         If (Shift And vbCtrlMask) = vbCtrlMask Then
            m_tI(i).bSelected = Not (m_tI(i).bSelected)
         ElseIf (Shift And vbShiftMask) = vbShiftMask Then
            If (iItem > m_iSelRow) Then
               For j = 1 To m_iIconCount
                  m_tI(j).bSelected = (j <= iItem) And (j >= m_iSelRow)
               Next j
            Else
               For j = 1 To m_iIconCount
                  m_tI(j).bSelected = (j >= iItem) And (j <= m_iSelRow)
               Next j
            End If
         Else
            For j = 1 To m_iIconCount
               m_tI(j).bSelected = (j = i)
            Next j
         End If
         m_iSelRow = i
         Draw
      End If
   End If
End Sub

Public Sub Create()
Dim tR As RECT
   m_lWidth = UserControl.ScaleWidth \ Screen.TwipsPerPixelY
   m_lHeight = UserControl.ScaleHeight \ Screen.TwipsPerPixelY
   Set m_cScroll = New cScrollBars
   m_cScroll.Create UserControl.hWnd
   m_cScroll.SmallChange(efsVertical) = 8
   DrawText UserControl.hdc, "Xg", 2, tR, DT_CALCRECT
   m_lTextHeight = tR.bottom - tR.top

End Sub

Public Sub Init(ByRef cI As cFileIcon)
Dim i As Long
   Clear
   If (cI.ImageCount > 0) Then
      m_iIconCount = cI.ImageCount
      ReDim m_tI(1 To m_iIconCount) As tIconInfo
      For i = 1 To m_iIconCount
         With m_tI(i)
            .hIcon = cI.IconHandle(UserControl.hdc, i)
            .lColours = cI.ImageColourCount(i)
            .lWidth = cI.ImageWidth(i)
            .lHeight = cI.ImageHeight(i)
         End With
      Next i
      Draw
   End If
End Sub

Private Sub pSetScroll()
Dim i As Long
Dim lHeight As Long
Dim lProportion As Long

If Not (m_cScroll Is Nothing) Then 'elh 2/20/99 if
  For i = 1 To m_iIconCount
    If (i > 1) Then
      m_tI(i).lStartY = lHeight
    End If
    lHeight = lHeight + m_lTextHeight + m_tI(i).lHeight + 6
  Next i
  If (lHeight > m_lHeight) Then
    m_cScroll.Max(efsVertical) = lHeight - m_lHeight
    lProportion = lHeight \ (lHeight - m_lHeight) + 1
    m_cScroll.LargeChange(efsVertical) = lProportion * m_lHeight
    If Not (m_cScroll.Visible(efsVertical)) Then
      m_cScroll.Visible(efsVertical) = True
    End If
  Else
    If (m_cScroll.Visible(efsVertical)) Then
      m_cScroll.Visible(efsVertical) = False
    End If
  End If
End If
End Sub

Public Sub Draw()
   pSetScroll
   Render
End Sub

Private Sub Render()
Dim i As Long
Dim lStartY As Long
Dim tR As RECT
Dim tTR As RECT
Dim sText As String
Dim hBr As Long
Dim bSelFlag As Boolean
Dim lHDC As Long

   If Not (m_cScroll Is Nothing) Then 'elh 2/20/99 if
      If (m_cScroll.Visible(efsVertical)) Then
         lStartY = -m_cScroll.Value(efsVertical)
      End If
   End If
   tR.right = m_lWidth
   tR.bottom = m_lHeight
   lHDC = UserControl.hdc
   hBr = GetSysColorBrush(vbWindowBackground And &H1F&)
   FillRect lHDC, tR, hBr
   DeleteObject hBr
   hBr = GetSysColorBrush(vbHighlight And &H1F&)
   
   tR.right = m_lWidth
   For i = 1 To m_iIconCount
      tR.top = m_tI(i).lStartY + lStartY
      tR.bottom = tR.top + m_lTextHeight + m_tI(i).lHeight + 6
      If ((tR.bottom) > 0) And (tR.top < m_lHeight) Then
         If (tR.top > m_lHeight) Then
            Exit For
         End If
         If (bSelFlag) And Not (m_tI(i).bSelected) Then
            SetBkMode lHDC, TRANSPARENT
            SetBkColor lHDC, GetSysColor(vbWindowBackground And &H1F&)
            SetTextColor lHDC, GetSysColor(vbWindowText And &H1F&)
         End If
         LSet tTR = tR
         If (m_tI(i).bSelected) Then
            FillRect lHDC, tTR, hBr
         End If
         tTR.right = tTR.right - 2
         tTR.left = 2 + (tTR.right - 4 - m_tI(i).lWidth) \ 2
         DrawIconEx lHDC, tTR.left, tTR.top + 2, m_tI(i).hIcon, m_tI(i).lWidth, m_tI(i).lHeight, 0, 0, DI_NORMAL
         
         LSet tTR = tR
         tTR.left = 2
         tTR.right = tTR.right - 2
         tTR.top = tTR.top + 2 + m_tI(i).lHeight + 2
         tTR.bottom = tTR.top + m_lTextHeight
         sText = m_tI(i).lWidth & " x " & m_tI(i).lHeight
         If (m_tI(i).lColours <= 256) Then
            sText = sText & ", " & m_tI(i).lColours & " colours"
         Else
            sText = sText & ", millions of colours"
         End If
         If (m_tI(i).bSelected) Then
            SetBkMode lHDC, OPAQUE
            SetBkColor lHDC, GetSysColor(vbHighlight And &H1F&)
            SetTextColor lHDC, GetSysColor(vbHighlightText And &H1F&)
            bSelFlag = True
         End If
         DrawText lHDC, sText, Len(sText), tTR, DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_CENTER
         If m_bInFocus And i = m_iSelRow Then
            LSet tTR = tR
            tTR.left = tTR.left + 1
            DrawFocusRect lHDC, tTR
         End If
      End If
   Next i
   
   DeleteObject hBr
   If (bSelFlag) Then
      SetBkMode lHDC, TRANSPARENT
      SetBkColor lHDC, GetSysColor(vbWindowBackground And &H1F&)
      SetTextColor lHDC, GetSysColor(vbWindowText And &H1F&)
   End If
End Sub
Public Sub GotFocus()
   m_bInFocus = True
   Draw
End Sub
Public Sub LostFocus()
   m_bInFocus = False
   Draw
End Sub
Public Sub Clear()
Dim i As Long
   For i = 1 To m_iIconCount
      DestroyIcon m_tI(i).hIcon
   Next i
   Erase m_tI
   m_iIconCount = 0
End Sub

Private Sub UserControl_GotFocus()
   GotFocus
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyDown KeyCode, Shift
End Sub

Private Sub UserControl_LostFocus()
   LostFocus
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
   If (Button And vbLeftButton) = vbLeftButton Then
      MouseDown Button, Shift, X, Y
   End If
End Sub

Private Sub UserControl_Paint()
   Draw
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Create
End Sub

Private Sub UserControl_Resize()
   m_lWidth = UserControl.ScaleWidth \ Screen.TwipsPerPixelY
   m_lHeight = UserControl.ScaleHeight \ Screen.TwipsPerPixelY
   Draw
End Sub

Private Sub UserControl_Terminate()
   Clear
   Set m_cScroll = Nothing
End Sub

Private Sub m_cScroll_Change(eBar As EFSScrollBarConstants)
   m_cScroll_Scroll eBar
End Sub

Private Sub m_cScroll_Scroll(eBar As EFSScrollBarConstants)
   Render
End Sub


