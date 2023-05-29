VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   4320
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Declarations
    Dim X As Single
    
    'Processing/Output
    For X = 32 To 127
        Print X; Chr$(X),
        If X Mod 5 = 1 Then
            Print
        End If
    Next X

End Sub
