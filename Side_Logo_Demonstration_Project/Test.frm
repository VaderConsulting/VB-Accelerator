VERSION 5.00
Begin VB.Form frmTest 
   Appearance      =   0  'Flat
   Caption         =   "Test SideBar Logo - Vertical Text"
   ClientHeight    =   3825
   ClientLeft      =   4365
   ClientTop       =   2625
   ClientWidth     =   6585
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   6585
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   9375
      Left            =   0
      ScaleHeight     =   9375
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cL As New cLogo
Private Sub Form_Load()
    cL.DrawingObject = picLogo
    cL.Caption = "Steve McMahon"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picLogo.Height = Me.ScaleHeight
    On Error GoTo 0
    cL.Draw
End Sub

