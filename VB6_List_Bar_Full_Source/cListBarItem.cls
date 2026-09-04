VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cListBarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private m_lBarID As Long
Private m_lBarPosition As Long
Private m_lItemID As Long
Private m_lItemPosition As Long

Friend Function fInit(ByVal hWnd As Long, ByVal lBarID As Long, ByVal lItemID As Long)
   m_hWnd = hWnd
   m_lBarID = lBarID
   m_lItemID = lItemID
End Function

Private Function pbVerify(ByRef ctlThis As vbalListBar) As Boolean
Dim lPtr As Long
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         m_lBarPosition = ctlThis.fIsBar(m_lBarID)
         If Not (m_lBarPosition = 0) Then
            m_lItemPosition = ctlThis.fIsItem(m_lBarPosition, m_lItemID)
            If Not (m_lItemPosition = 0) Then
               pbVerify = True
            Else
               gErr 3, "cListBarItem" ' Item does not exist
            End If
         Else
            gErr 2, "cListBarItem" ' Bar does not exist
         End If
      Else
         gErr 1, "cListBarItem" ' object has been unloaded
      End If
   Else
      gErr 1, "cListBarItem" ' object has been unloaded
   End If
End Function
Public Property Get HelpText() As String
Attribute HelpText.VB_Description = "Gets/sets the Help Text which will be shown when the user hovers over the item within the bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      HelpText = ctl.fItemHelpText(m_lBarPosition, m_lItemPosition)
   End If
End Property
Public Property Let HelpText(ByVal sHelpText As String)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fItemHelpText(m_lBarPosition, m_lItemPosition) = sHelpText
   End If
End Property
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Gets/sets the caption of the item."
Attribute Caption.VB_UserMemId = 0
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Caption = ctl.fItemCaption(m_lBarPosition, m_lItemPosition)
   End If
End Property
Public Property Let Caption(ByVal sCaption As String)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fItemCaption(m_lBarPosition, m_lItemPosition) = sCaption
   End If
End Property
Public Property Get Key() As String
Attribute Key.VB_Description = "Gets/sets the Key associated with the item."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Key = ctl.fItemKey(m_lBarPosition, m_lItemPosition)
   End If
End Property
Public Property Let Key(ByVal sKey As String)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fItemKey(m_lBarPosition, m_lItemPosition) = sKey
   End If
End Property
Public Property Get Tag() As String
Attribute Tag.VB_Description = "Gets/sets a string value which can be used to associate data with the item."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Tag = ctl.fItemTag(m_lBarPosition, m_lItemPosition)
   End If
End Property
Public Property Let Tag(ByVal sTag As String)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fItemTag(m_lBarPosition, m_lItemPosition) = sTag
   End If
End Property
Public Property Get itemData() As Long
Attribute itemData.VB_Description = "Gets/sets a long value associated with the item."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      itemData = ctl.fItemItemData(m_lBarPosition, m_lItemPosition)
   End If
End Property
Public Property Let itemData(ByVal lItemData As Long)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fItemItemData(m_lBarPosition, m_lItemPosition) = lItemData
   End If
End Property
Public Property Get IconIndex() As Long
Attribute IconIndex.VB_Description = "Gets/sets the 0 based index of the icon within the ImageList to show for this item."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      IconIndex = ctl.fItemIconIndex(m_lBarPosition, m_lItemPosition)
   End If
End Property
Public Property Let IconIndex(ByVal lIconIndex As Long)
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fItemIconIndex(m_lBarPosition, m_lItemPosition) = lIconIndex
   End If
End Property
Public Sub BeginEdit()
Attribute BeginEdit.VB_Description = "Initiates in-place editing of the item within the bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBeginEdit m_lBarPosition, m_lItemPosition
   End If
End Sub
Public Property Get Left() As Long
Attribute Left.VB_Description = "Gets the Left position of the item within the bar."
Dim ctl As vbalListBar
Dim tR As RECT
   If pbVerify(ctl) Then
      ctl.fItemGetRect m_lBarPosition, m_lItemPosition, tR
      Left = tR.Left
   End If
End Property
Public Property Get Top() As Long
Attribute Top.VB_Description = "Gets the Top position of the item within the bar."
Dim ctl As vbalListBar
Dim tR As RECT
   If pbVerify(ctl) Then
      ctl.fItemGetRect m_lBarPosition, m_lItemPosition, tR
      Top = tR.Top
   End If
End Property
Public Property Get Width() As Long
Attribute Width.VB_Description = "Gets the Width of the item within the bar."
Dim ctl As vbalListBar
Dim tR As RECT
   If pbVerify(ctl) Then
      ctl.fItemGetRect m_lBarPosition, m_lItemPosition, tR
      Width = (tR.Right - tR.Left)
   End If
End Property
Public Property Get Height() As Long
Attribute Height.VB_Description = "Gets the Height of the item."
Dim ctl As vbalListBar
Dim tR As RECT
   If pbVerify(ctl) Then
      ctl.fItemGetRect m_lBarPosition, m_lItemPosition, tR
      Height = (tR.Bottom - tR.Top)
   End If
End Property
Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Gets whether the item was the last selected item in the bar or not."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      Selected = (ctl.fBarSelectedItemID(m_lBarPosition) = m_lItemID)
   End If
End Property
Public Sub SelectItem()
Attribute SelectItem.VB_Description = "Selects this item within the bar."
Dim ctl As vbalListBar
   If pbVerify(ctl) Then
      ctl.fBarItemSelect m_lBarPosition, m_lItemPosition
   End If
End Sub

