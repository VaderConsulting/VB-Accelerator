VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{FE1D1F8B-EC4B-11D3-B06C-00500427A693}#1.0#0"; "vbalLBar6.ocx"
Begin VB.Form frmTestListBar 
   Caption         =   "List Bar Control Tester"
   ClientHeight    =   5130
   ClientLeft      =   2580
   ClientTop       =   2160
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTesLBar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   5715
   Begin vbalIml6.vbalImageList ilsIcons16 
      Left            =   2160
      Top             =   2040
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   12628
      Images          =   "frmTesLBar.frx":1272
      Version         =   131072
      KeyCount        =   11
      Keys            =   "ÿSystemÿExplorerÿFavouritesÿCalendarÿNetwork NeighbourhoodÿHistoryÿInternet ExplorerÿMailÿNewsÿChannels"
   End
   Begin vbalIml6.vbalImageList ilsIcons32 
      Left            =   2160
      Top             =   1380
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   24
      Size            =   48532
      Images          =   "frmTesLBar.frx":43E6
      Version         =   131072
      KeyCount        =   11
      Keys            =   "FindÿSystemÿExplorerÿFavouritesÿCalendarÿNetwork NeighbourhoodÿHistoryÿInternet ExplorerÿMailÿNewsÿChannels"
   End
   Begin vbalLbar6.vbalListBar vbalListBar1 
      Height          =   4635
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   8176
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picStatus 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   5475
      TabIndex        =   4
      Top             =   4800
      Width           =   5475
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "vbAccelerator ListBar Control"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2100
      End
   End
   Begin VB.PictureBox picMainFrame 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   3060
      ScaleHeight     =   3615
      ScaleWidth      =   2295
      TabIndex        =   2
      Top             =   720
      Width           =   2295
      Begin VB.CommandButton cmdClearBar 
         BackColor       =   &H80000005&
         Caption         =   "&Clear Bar..."
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   3300
         Width           =   1455
      End
      Begin VB.CheckBox chkOfficeXpStyle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Office &Xp Style"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Top             =   1740
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H80000005&
         Caption         =   "&Add to Last Bar"
         Height          =   315
         Left            =   360
         TabIndex        =   16
         Top             =   2940
         Width           =   1455
      End
      Begin VB.PictureBox picCont 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   60
         ScaleHeight     =   795
         ScaleWidth      =   2055
         TabIndex        =   12
         Top             =   2160
         Width           =   2055
         Begin VB.OptionButton optIconSize 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Small Icons"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   15
            Top             =   480
            Width           =   1755
         End
         Begin VB.OptionButton optIconSize 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Large Icons"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Icon Size:"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox picWidth 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   180
         ScaleHeight     =   735
         ScaleWidth      =   1935
         TabIndex        =   8
         Top             =   300
         Width           =   1935
         Begin VB.OptionButton optWidth 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Regular"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optWidth 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Posh Spice"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   10
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton optWidth 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Who ate all the pies?"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkBackground 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Background Bitmap"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Visual Settings:"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   1260
         Width           =   1815
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "ListBar Width:"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   1815
      End
   End
   Begin VB.Label lblContextDetail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ListBar Today"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   3060
      TabIndex        =   1
      Top             =   180
      Width           =   1605
   End
   Begin VB.Label lblContext 
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New Window..."
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Index           =   2
      End
   End
   Begin VB.Menu mnuPopupTOP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "&Open <i>..."
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Open in New &Window"
         Index           =   1
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Advanced &Find..."
         Index           =   2
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Re&move from <b>"
         Index           =   4
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "&Rename Shortcut"
         Index           =   5
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Propert&ies"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmTestListBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bInhibitOptionClick As Boolean

' -----------------------------------------------------------------------
' For setting up a thin border on a picture box control:
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40

Private Function ThinBorder(ByVal lhWnd As Long, ByVal bState As Boolean)
Dim lS As Long

   lS = GetWindowLong(lhWnd, GWL_EXSTYLE)
   If Not (bState) Then
      lS = lS Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
   Else
      lS = lS Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
   End If
   SetWindowLong lhWnd, GWL_EXSTYLE, lS
   SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Function
' -----------------------------------------------------------------------

Public Sub TestCollections()
Dim i As Long
Dim j As Long
   ' This function just enumerates through the various items
   ' in the control to check they all work ok.
   ' Note that For..Each is not support in this release.
   With vbalListBar1
      For i = 1 To .Bars.Count
         With .Bars(i)
            Debug.Print "Bar:"; i & ";" & .Caption
            For j = 1 To .Items.Count
               Debug.Print "   Item:" & j & ";" & .Items(j).Caption
            Next j
         End With
      Next i
   End With
End Sub


Private Sub chkBackground_Click()
Dim i As Long
   If chkBackground.Value = vbChecked Then
      vbalListBar1.BackgroundPicture = App.Path & "\PARCHMTDL.JPG"
      For i = 1 To vbalListBar1.Bars.Count
         vbalListBar1.Bars(i).ForeColor = vbWindowText
      Next i
   Else
      vbalListBar1.BackgroundPicture = ""
      For i = 1 To vbalListBar1.Bars.Count
         vbalListBar1.Bars(i).ForeColor = vbWindowBackground
      Next i
   End If
End Sub

Private Sub chkOfficeXPStyle_Click()
Dim barX As cListBar
   If Not (m_bInhibitOptionClick) Then
      Set barX = vbalListBar1.SelectedBar
      If Not (barX Is Nothing) Then
         If (chkOfficeXpStyle.Value = vbChecked) Then
            barX.BackColor = vbButtonFace
            barX.ForeColor = vbWindowText
            barX.OfficeXpStyle = True
         Else
            barX.BackColor = vbButtonShadow
            barX.ForeColor = vb3DHighlight
            barX.OfficeXpStyle = False
         End If
      End If
   End If
End Sub

Private Sub cmdAdd_Click()
Dim sRndCap As String
Dim iRndIcon As Long
Dim i As Long
   For i = 1 To Rnd * 72 + 1
      If Rnd > 0.7 Then
         sRndCap = sRndCap & " "
      Else
         sRndCap = sRndCap & Chr$(Rnd * 26 + 65)
      End If
   Next i
   iRndIcon = (Rnd * ilsIcons32.ImageCount + 1) - 1
   With vbalListBar1.Bars(vbalListBar1.Bars.Count)
      .Items.Add "Test" & .Items.Count, , sRndCap, iRndIcon
   End With
End Sub

Private Sub cmdClearBar_Click()
Dim barX As cListBar
   Set barX = vbalListBar1.SelectedBar
   If Not (barX Is Nothing) Then
      If (vbYes = MsgBox("Are you sure you want to clear the bar " & barX.Caption & "?", vbYesNo Or vbQuestion)) Then
         vbalListBar1.SelectedBar.Items.Clear
      End If
   End If
End Sub

Private Sub Form_Load()
Dim barX As cListBar
Dim itmX As cListBarItem
Dim i As Long

   ' This is a good effect for 98/2000 style apps:
   ThinBorder picMainFrame.hwnd, True
   ThinBorder picStatus.hwnd, True
   
   ' Set up the ListBar:
   With vbalListBar1
      
      ' vbAccelerator ImageLists (you can happily share these with other controls
      ' in your app):
      .ImageList(evlbLargeIcon) = ilsIcons32
      .ImageList(evlbSmallIcon) = ilsIcons16
      
      
      ' Add a bar and add some items to it:
      Set barX = .Bars.Add("OUTLOOK", , "Outlook Shortcuts")
      For i = 1 To ilsIcons32.ImageCount
         Set itmX = barX.Items.Add(.Bars.Count & "Item" & i, , ilsIcons32.ItemKey(i), i - 1)
         If i = 2 Then
            ' Demonstrate the InfoTips capabilities:
            itmX.HelpText = "The System Tab Contains commands for working with your computer's disk drives and interfaces"
         End If
      Next i
      
      ' Add a bar and add less items to it:
      Set barX = .Bars.Add("CUSTOM", , "My Shortcuts")
      For i = 2 To ilsIcons32.ImageCount Step 3
         barX.Items.Add .Bars.Count & "Item" & i, , ilsIcons32.ItemKey(i), i - 1
      Next i
            
      ' 1 item in this bar:
      Set barX = .Bars.Add("OTHER", , "Other Shortcuts")
      barX.Items.Add .Bars.Count & "Item1", , ilsIcons32.ItemKey(5), 4
      
      ' Check no items in a bar:
      .Bars.Add "USER1", , "Empty Bar"
               
   End With
   
   TestCollections
   
End Sub

Private Sub Form_Resize()
Dim lL As Long
Dim lT As Long
On Error Resume Next
   vbalListBar1.Move 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY, vbalListBar1.Width, Me.ScaleHeight - picStatus.Height - 6 * Screen.TwipsPerPixelY
   lL = vbalListBar1.Left + vbalListBar1.Width + 3 * Screen.TwipsPerPixelX
   lblContext.Move lL, vbalListBar1.Top, Me.ScaleWidth - lL - 2 * Screen.TwipsPerPixelX
   lblContextDetail.Move lblContext.Left + 2 * Screen.TwipsPerPixelX, lblContext.Top + (lblContext.Height - lblContextDetail.Height) \ 2
   lT = lblContext.Top + lblContext.Height + 3 * Screen.TwipsPerPixelY
   picMainFrame.Move lblContext.Left, lT, lblContext.Width, Me.ScaleHeight - lT - 4 * Screen.TwipsPerPixelY - picStatus.Height
   picStatus.Move 2 * Screen.TwipsPerPixelX, vbalListBar1.Top + vbalListBar1.Height + 2 * Screen.TwipsPerPixelY, Me.ScaleWidth - 4 * Screen.TwipsPerPixelX
   lblStatus.Move 2 * Screen.TwipsPerPixelX, (picStatus.ScaleHeight - lblStatus.Height) \ 2
End Sub

Private Sub mnuFile_Click(Index As Integer)
Dim lL As Long, lT As Long
   Select Case Index
   Case 0
      
      ' Never forget to test what happens when multiple instances
      ' of your control run from the same project!!!
      Dim f As New frmTestListBar
      f.Show
      
      
      ' A sort of Cascade mechanism:
      If f.Left = Me.Left And f.Top = Me.Top Then
         lL = f.Left + (48 * Screen.TwipsPerPixelX)
         ' Should really use SystemParametersInfo call to get desktop area
         ' here
         If lL + f.Width < Screen.Width Then
            lL = 48 * Screen.TwipsPerPixelY
         End If
         lT = f.Top + (48 * Screen.TwipsPerPixelY)
         If lT + f.Height < Screen.Height Then
            lT = 48 * Screen.TwipsPerPixelX
         End If
         f.Move lL, lT
      End If
   
   Case 2
      ' Game Over
      Unload Me
      
   End Select
End Sub

Private Sub mnuPopup_Click(Index As Integer)
   Select Case Index
   Case 0
      vbalListBar1.SelectedBar.Items(mnuPopupTOP.Tag).SelectItem
   Case 1
      Dim f As New frmTestListBar
      f.Show
   Case 2
      MsgBox "Show Advanced Find Dialog here.", vbInformation
   Case 4
      ' Remove
      If MsgBox("Are you sure you want to remove the item " & vbalListBar1.SelectedBar.Items(mnuPopupTOP.Tag) & "?", vbYesNo Or vbQuestion) = vbYes Then
         vbalListBar1.SelectedBar.Items.Remove mnuPopupTOP.Tag
      End If
   Case 5
      ' Rename:
      vbalListBar1.SelectedBar.Items(mnuPopupTOP.Tag).BeginEdit
   Case 7
      MsgBox "Show Properties Dialog here.", vbInformation
   End Select
End Sub

Private Sub optIconSize_Click(Index As Integer)
   If Not (m_bInhibitOptionClick) Then
      If optIconSize(0).Value Then
         vbalListBar1.Bars(vbalListBar1.SelectedBar.Key).IconSize = evlbLargeIcon
      Else
         vbalListBar1.Bars(vbalListBar1.SelectedBar.Key).IconSize = evlbSmallIcon
      End If
   End If
End Sub

Private Sub optWidth_Click(Index As Integer)
   Select Case Index
   Case 0
      vbalListBar1.Width = 128 * Screen.TwipsPerPixelX
   Case 1
      vbalListBar1.Width = 32 * Screen.TwipsPerPixelX
   Case 2
      vbalListBar1.Width = 256 * Screen.TwipsPerPixelX
   End Select
   Form_Resize
End Sub





Private Sub vbalListBar1_BarClick(Bar As vbalLBar6.cListBar)
   ' check for new bar icon setting:
   m_bInhibitOptionClick = True
      If Bar.IconSize = evlbLargeIcon Then
         optIconSize(0).Value = True
      Else
         optIconSize(1).Value = True
      End If
      If Bar.OfficeXpStyle Then
         chkOfficeXpStyle.Value = vbChecked
      Else
         chkOfficeXpStyle.Value = vbUnchecked
      End If
   m_bInhibitOptionClick = False
End Sub

Private Sub vbalListBar1_ItemClick(Item As vbalLBar6.cListBarItem, Bar As vbalLBar6.cListBar)
   MsgBox "Clicked Item " & Item.Caption & vbCrLf & "In bar " & Bar.Caption, vbInformation
End Sub

Private Sub vbalListBar1_ItemEndEdit(Item As vbalLBar6.cListBarItem, sText As String, bCancel As Boolean)
Dim i As Long
   If Len(sText) < 2 Then
      MsgBox "Please enter a caption more than 1 character long.", vbInformation
      bCancel = True
   Else
      For i = 1 To vbalListBar1.SelectedBar.Items.Count
         If vbalListBar1.SelectedBar.Items(i) <> Item Then
            If sText = vbalListBar1.SelectedBar.Items(i).Caption Then
               If vbYes = MsgBox("There is already an item named '" & sText & "' in the ListBar." & vbCrLf & vbCrLf & "Are you sure you want to rename it?", vbYesNo Or vbQuestion) Then
               Else
                  bCancel = True
               End If
            End If
         End If
      Next i
   End If
End Sub

Private Sub vbalListBar1_ItemRightClick(Item As vbalLBar6.cListBarItem, Bar As vbalLBar6.cListBar, X As Single, Y As Single)
   mnuPopup(0).Caption = "&Open " & Item.Caption & "..."
   mnuPopup(4).Caption = "Re&move from " & Bar.Caption & "..."
   mnuPopupTOP.Tag = Item.Key
   Me.PopupMenu mnuPopupTOP, , X + vbalListBar1.Left, Y + vbalListBar1.Top
End Sub

