VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   6585
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim MyArray(255) As Byte
Dim I As Integer
Dim CRC As clsCRC32

    Set CRC = New clsCRC32
    For I = 0 To 16
        MyArray(I) = I
    Next I
    Debug.Print CRC.CalcCRC32(MyArray)

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Oct-27 12:17) 1 + 16 = 17 Lines
