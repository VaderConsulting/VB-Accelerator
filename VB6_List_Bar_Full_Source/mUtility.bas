Attribute VB_Name = "mUtility"
Option Explicit

Public Const gcObjectProp = "vbalListBar:ObjectPtr"

Private m_lBarID As Long
Private m_lItemID As Long


Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim objT As Object
   If Not (lPtr = 0) Then
      ' Turn the pointer into an illegal, uncounted interface
      CopyMemory objT, lPtr, 4
      ' Do NOT hit the End button here! You will crash!
      ' Assign to legal reference
      Set ObjectFromPtr = objT
      ' Still do NOT hit the End button here! You will still crash!
      ' Destroy the illegal reference
      CopyMemory objT, 0&, 4
   End If
End Property

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Property Get gNewBarID() As Long
   m_lBarID = m_lBarID + 1
   gNewBarID = m_lBarID
End Property
Public Property Get gNewItemID() As Long
   m_lItemID = m_lItemID + 1
   gNewItemID = m_lItemID
End Property
Public Sub gErr(ByVal lErrNum As Long, ByVal sSource As String)
Dim sDesc As String
Debug.Assert False
   Select Case lErrNum
   Case 1
      ' Cannot find owner object
      lErrNum = 364
      sDesc = "Object has been unloaded."
   Case 2
      ' Bar does not exist
      lErrNum = vbObjectError + 25001
      sDesc = "ListBar does not exist."
      
   Case 3
      ' Item does not exist
      lErrNum = vbObjectError + 25002
      sDesc = "ListItem does not exist."
      
   Case 4
      ' Invalid key: numeric
      lErrNum = 13
      sDesc = "Type Mismatch."
      
   Case 5
      ' Invalid Key: duplicate
      lErrNum = 457
      sDesc = "This key is already associated with an element of this collection."
   
   Case 6
      ' Subscript out of range
      lErrNum = 9
      sDesc = "Subscript out of range."
   
   Case Else
      Debug.Assert "Unexpected Error" = ""
   
   End Select
   
   
   Err.Raise lErrNum, App.EXEName & "." & sSource, sDesc
End Sub

