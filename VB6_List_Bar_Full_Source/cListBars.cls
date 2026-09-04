VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cListBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long

Friend Function fInit(ByVal hWnd As Long)
   m_hWnd = hWnd
End Function
Private Function pbVerify(ByRef ctlThis As vbalListBar) As Boolean
Dim lPtr As Long
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         pbVerify = True
      Else
         gErr 1, "cListBars"
      End If
   Else
      gErr 1, "cListBars"
   End If
End Function
Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the number of bars within the control."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Count = ctl.fBarCount
   End If
End Property
Public Property Get Item(Key As Variant) As cListBar
Attribute Item.VB_Description = "Gets a reference to a specific Bar."
Attribute Item.VB_UserMemId = 0
Dim ctl As vbalListBar
Dim lID As Long
   If pbVerify(ctl) Then
      lID = ctl.fBarIDForKey(Key)
      If lID > 0 Then
         Dim cL As cListBar
         Set cL = New cListBar
         cL.fInit m_hWnd, lID
         Set Item = cL
      End If
   End If
End Property
Public Function Add(Optional Key As Variant, Optional Before As Variant, Optional ByVal sCaption As String, Optional sHelpText As String) As cListBar
Attribute Add.VB_Description = "Adds or inserts a new bar into the control."
Dim ctl As vbalListBar
Dim lID As Long
   If pbVerify(ctl) Then
      lID = ctl.fAddBar(Key, Before, sCaption, sHelpText)
      If lID > 0 Then
         Dim cL As cListBar
         Set cL = New cListBar
         cL.fInit m_hWnd, lID
         Set Add = cL
      End If
   End If
End Function
Public Sub Remove(Key As Variant)
Attribute Remove.VB_Description = "Removes a bar from the control."
Dim ctl As vbalListBar
Dim lID As Long
   If pbVerify(ctl) Then
      ctl.fRemoveBar Key
   End If
End Sub
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all bars from the control."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fClear
   End If
End Sub
