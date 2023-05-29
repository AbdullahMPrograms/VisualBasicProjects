VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblPassword 
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Password:"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Insert Any Alphanumeric Character:"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Insert Any Character:"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Insert Digits or Backspace:"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Insert Letter in Any Case:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Insert Lower Case Letters Only:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Password As String

Private Sub Text1_KeyPress(KeyAscii As Integer)

Dim Char As String

Char = Chr$(KeyAscii)
If Char < "a" Or Char > "z" Then
    KeyAscii = 0
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

Dim Char As String

Char = Chr$(KeyAscii)
If (Char < "a" Or Char > "z") And (Char < "A" Or Char > "Z") Then
    KeyAscii = 0
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

Dim Char As String

Char = Chr$(KeyAscii)
If KeyAscii <> 8 And (Char < "0" Or Char > "9") Then
    KeyAscii = 0
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

Dim Char As String

Char = Chr$(KeyAscii)
If Char >= "a" And Char <= "z" Then
    KeyAscii = KeyAscii - 32

End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)

Dim Char As String

Char = Chr$(KeyAscii)
If (Char <= "a" Or Char >= "z") And (Char < "0" Or Char > "9") And (Char <= "A" Or Char >= "Z") Then
    KeyAscii = 0

End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

Dim Char As String
Dim K As Integer


Char = Chr$(KeyAscii)
If KeyAscii <> 8 And (Char < "0" Or Char > "9") Then
    KeyAscii = 0
Else
    If KeyAscii = 8 Then
        Password = Left$(Password, Len(Password) - 1)
    Else
        KeyAscii = 42
        Password = Password + Char
    End If
End If

lblPassword.Caption = Password

End Sub
