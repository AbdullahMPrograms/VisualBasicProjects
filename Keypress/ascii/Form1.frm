VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblKey 
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)

 lblKey.Caption = "You pressed ASCII value " & Str$(KeyAscii) & " or " & Chr$(KeyAscii)

End Sub

