VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cListBarItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private m_lID As Long
Private m_lPosition As Long

Friend Function fInit(ByVal hWnd As Long, ByVal lID As Long)
   m_hWnd = hWnd
   m_lID = lID
End Function

Private Function pbVerify(ByRef ctlThis As vbalListBar) As Boolean
Dim lPtr As Long
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         m_lPosition = ctlThis.fIsBar(m_lID)
         If Not (m_lPosition = 0) Then
            pbVerify = True
         Else
            gErr 2, "cListBarItems" ' Bar does not exist
         End If
      Else
         gErr 1, "cListBarItems" ' object has been unloaded
      End If
   Else
      gErr 1, "cListBarItems" ' object has been unloaded
   End If
End Function

Public Property Get SelectedItem() As cListBarItem
Attribute SelectedItem.VB_Description = "Gets the last selected item within the bar."
Dim ctl As vbalListBar
Dim lItemID As Long
Dim lPos As Long
   If pbVerify(ctl) Then
      lPos = ctl.fItemSelectedID(m_lPosition)
      If lPos > 0 Then
         lItemID = ctl.fItemID(m_lPosition, lPos)
         Dim cLBI As cListBarItem
         Set cLBI = New cListBarItem
         cLBI.fInit m_hWnd, m_lID, lItemID
         Set SelectedItem = cLBI
      End If
   End If
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items within the Bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Count = ctl.fItemCount(m_lPosition)
   End If
End Property
Public Property Get Item(Key As Variant) As cListBarItem
Attribute Item.VB_Description = "Gets a specific item within the bar."
Attribute Item.VB_UserMemId = 0
Dim ctl As vbalListBar
Dim lItemID As Long
Dim lPos As Long
   If pbVerify(ctl) Then
      lPos = ctl.fItemIndex(m_lPosition, Key)
      If lPos > 0 Then
         lItemID = ctl.fItemID(m_lPosition, lPos)
         Dim cLBI As cListBarItem
         Set cLBI = New cListBarItem
         cLBI.fInit m_hWnd, m_lID, lItemID
         Set Item = cLBI
      End If
   End If
End Property
Public Function Add(Optional Key As Variant, Optional Before As Variant, Optional ByVal sCaption As String, Optional lIconIndex As Long = -1) As cListBarItem
Attribute Add.VB_Description = "Adds or Inserts a new Item into the ListBar."
Dim ctl As vbalListBar
Dim lID As Long
   If pbVerify(ctl) Then
      lID = ctl.fAddItem(m_lPosition, Key, Before, sCaption, lIconIndex)
      If lID > 0 Then
         Dim cLBI As cListBarItem
         Set cLBI = New cListBarItem
         cLBI.fInit m_hWnd, m_lID, lID
         Set Add = cLBI
      End If
   End If
End Function
Public Sub Remove(Key As Variant)
Attribute Remove.VB_Description = "Removes an item within the bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fRemoveItem m_lPosition, Key
   End If
End Sub
Public Sub Clear()
Attribute Clear.VB_Description = "Clears all items in this bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fClearBar m_lPosition
   End If
End Sub
