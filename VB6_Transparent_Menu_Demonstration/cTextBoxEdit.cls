VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTextBoxEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageString Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_COMMAND = &H111
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const EM_UNDO = &HC7
Private Const EM_CANUNDO = &HC6
Private Const EM_REPLACESEL = &HC2
Private Declare Function IsClipboardFormatAvailable Lib "USER32" _
      (ByVal wFormat As Long) As Long
Private Const CF_TEXT = 1
Private Const CF_UNICODETEXT = 13
Private Const CF_OEMTEXT = 7

Private m_txtThis As TextBox

Public Property Let TextBox(ByRef txtThis As TextBox)
   Set m_txtThis = txtThis
End Property
Public Sub Cut()
   SendMessageLong m_txtThis.hWnd, WM_CUT, 0, 0
End Sub
Public Sub Copy()
   SendMessageLong m_txtThis.hWnd, WM_COPY, 0, 0
End Sub
Public Sub Paste()
   SendMessageLong m_txtThis.hWnd, WM_PASTE, 0, 0
End Sub
Public Sub Undo()
   If (SendMessageLong(m_txtThis.hWnd, EM_CANUNDO, 0, 0) <> 0) Then
      SendMessageLong m_txtThis.hWnd, EM_UNDO, 0, 0
   End If
End Sub
Public Property Get CanCut() As Boolean
   CanCut = (Not (m_txtThis.Locked) And m_txtThis.SelLength > 0)
End Property
Public Property Get CanCopy() As Boolean
   CanCopy = (m_txtThis.SelLength > 0)
End Property
Public Property Get CanPaste() As Boolean
   If IsClipboardFormatAvailable(CF_TEXT) Then
      CanPaste = True
   ElseIf IsClipboardFormatAvailable(CF_UNICODETEXT) Then
      CanPaste = True
   ElseIf IsClipboardFormatAvailable(CF_OEMTEXT) Then
      CanPaste = True
   End If
End Property
Public Property Get CanUndo() As Boolean
   CanUndo = (SendMessageLong(m_txtThis.hWnd, EM_CANUNDO, 0, 0) <> 0)
End Property
Public Sub ReplaceSelection(ByRef sText As String, Optional ByVal bAllowUndo = True)
   Dim lR As Long
   If (m_txtThis.SelLength > 0) Then
      lR = Abs(bAllowUndo)
      SendMessageString m_txtThis.hWnd, EM_REPLACESEL, lR, sText
   End If
End Sub
Public Sub Delete(Optional ByVal bAllowUndo = True)
   Dim lR As Long
   SendMessageString m_txtThis.hWnd, EM_REPLACESEL, lR, vbNullChar
End Sub


