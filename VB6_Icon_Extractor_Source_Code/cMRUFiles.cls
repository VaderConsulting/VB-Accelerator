VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMRUFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFilename As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Const MAX_PATH = 260

Private m_sPath() As String
Private m_iCount As Long
Private m_iMaxCount As Long

Public Function Load(cR As cRegistry) As Boolean
   Load = ReadProperties(cR)
End Function
Public Function Save(cR As cRegistry) As Boolean
   WriteProperties cR
End Function

Public Property Get Count() As Long
   Count = m_iCount
End Property

Public Property Get MaxCount() As Long
   MaxCount = m_iMaxCount
End Property
Public Property Let MaxCount(ByVal lMaxCount As Long)
Dim i As Long
   If lMaxCount = m_iMaxCount Then
      Exit Property
   End If
   If lMaxCount < m_iCount Then
      m_iCount = lMaxCount
   End If
   m_iMaxCount = lMaxCount
   ReDim Preserve m_sPath(1 To m_iMaxCount) As String
   
End Property
Private Sub pGetFile(ByVal nIndex As Long, ByRef sPath As String, ByRef sFile As String, ByRef sExt As String)
Dim sJunk As String
Dim iPos As Long
Dim i As Long
Dim iExtPos As Long
   sPath = m_sPath(nIndex)
   iPos = InStr(sPath, vbNullChar)
   If iPos > 1 Then
      sPath = left$(sPath, iPos - 1)
   End If
   For i = Len(sPath) To 1 Step -1
      If sExt = "" Then
         If Mid$(sPath, i, 1) = "." Then
            sExt = right$(sPath, i + 1)
            iExtPos = i
         End If
      End If
      If sFile = "" Then
         If Mid$(sPath, i, 1) = "\" Then
            If iExtPos > 0 Then
               sFile = Mid$(sPath, i + 1, iExtPos - i - 1)
            Else
               sFile = Mid$(sPath, i + 1)
            End If
            sPath = left$(sPath, i - 1)
            Exit For
         End If
      End If
   Next i
End Sub
Public Property Get File(ByVal nIndex As Long) As String
Dim sPath As String
Dim sFile As String
Dim sExt As String
Dim iPos As Long
   pGetFile nIndex, sPath, sFile, sExt
   File = sFile
End Property
Public Property Get Folder(ByVal nIndex As Long) As String
Dim sPath As String
Dim sFile As String
Dim sExt As String
Dim iPos As Long
   pGetFile nIndex, sPath, sFile, sExt
   Folder = sPath
End Property
Public Property Get Extension(ByVal nIndex As Long) As String
Dim sPath As String
Dim sFile As String
Dim sExt As String
Dim iPos As Long
   pGetFile nIndex, sPath, sFile, sExt
   Extension = sExt
End Property
Public Property Get Path(ByVal nIndex As Long) As String
   Path = m_sPath(nIndex)
End Property
Public Sub Delete(ByVal nIndex As Long)
Dim i As Long
   If nIndex > 0 And nIndex <= m_iMaxCount Then
      For i = nIndex To m_iMaxCount - 1
         m_sPath(i) = m_sPath(i + 1)
         If i > m_iCount Then
            m_sPath(i) = ""
         End If
      Next i
      m_sPath(m_iMaxCount) = ""
      m_iCount = m_iCount - 1
   End If
End Sub
Private Function ShortPath(ByVal sPath As String) As String
Dim sShortPath As String
Dim iPos As Long
   sShortPath = String$(MAX_PATH * 2 + 1, 0)
   GetShortPathName sPath, sShortPath, MAX_PATH * 2
   iPos = InStr(sShortPath, vbNullChar)
   If iPos > 1 Then
      ShortPath = left$(sShortPath, iPos - 1)
   Else
      ShortPath = sShortPath
   End If
End Function
Public Function ItemIndex(ByVal sPath As String) As Long
Dim sComp As String
Dim sShortPath As String
Dim iPos As Long
Dim i As Long

   sShortPath = ShortPath(sPath)
   ' Have we already got it?
   For i = 1 To m_iCount
      If UCase(ShortPath(m_sPath(i))) = UCase(sShortPath) Then
         ItemIndex = i
         Exit For
      End If
   Next i
   
End Function
Public Sub Add(ByVal sPath As String)
Dim sComp As String
Dim sShortPath As String
Dim iPos As Long
Dim i As Long
Dim lAlready As Long
   
   sShortPath = ShortPath(sPath)
   lAlready = ItemIndex(sPath)
   If lAlready = 0 Then
      ' Add the item:
      For i = m_iMaxCount - 1 To 1 Step -1
         m_sPath(i + 1) = m_sPath(i)
      Next i
   Else
      ' Swap lAlready to position 1:
      For i = lAlready - 1 To 1 Step -1
         m_sPath(i + 1) = m_sPath(i)
      Next i
   End If
   m_sPath(1) = sPath
   If lAlready = 0 Then
      m_iCount = m_iCount + 1
      If m_iCount > m_iMaxCount Then
         m_iCount = m_iMaxCount
      End If
   End If
End Sub

Private Sub Class_Initialize()
   pInit
End Sub
Private Sub pInit()
   MaxCount = 8
   m_iCount = 0
End Sub

Private Function ReadProperties(cR As cRegistry) As Boolean
Dim i As Long
Dim iCount As Long
   With cR
      .ValueKey = "MaxCount"
      .ValueType = REG_DWORD
      .Default = -1
      ReadProperties = (.KeyExists) And (.Value > 0)
      If .Value < 1 Then
         Exit Function
      End If
      MaxCount = .Value
      If m_iMaxCount > 0 Then
         .ValueKey = "Count"
         m_iCount = .Value
         If m_iCount > m_iMaxCount Then
            m_iCount = m_iMaxCount
         End If
         .ValueType = REG_SZ
         For i = 1 To m_iCount
            .ValueKey = "Path" & i
            m_sPath(i) = .Value
         Next i
      End If
   End With
End Function

Private Sub WriteProperties(cR As cRegistry)
Dim i As Long
Dim iCount As Long
   With cR
      .ValueKey = "MaxCount"
      .ValueType = REG_DWORD
      .Value = MaxCount
      .ValueKey = "Count"
      .Value = Count
      .ValueType = REG_SZ
      For i = 1 To Count
         .ValueKey = "Path" & i
         .Value = m_sPath(i)
      Next i
   End With
End Sub

