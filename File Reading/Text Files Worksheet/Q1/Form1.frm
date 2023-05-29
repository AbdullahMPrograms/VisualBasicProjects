VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   240
      ScaleHeight     =   7395
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   240
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   7920
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Declarations
Dim X As Integer
Dim Average As Single
Dim ItemName As String
Dim Price As Single


X = 0
picData.Cls
picData.Print "ItemName"; Tab(25); "Price"
picData.Print

'Input
Open "F:\ICS\File Reading\Text Files Worksheet\Prices.txt" For Input As #1
Do While Not EOF(1)
    X = X + 1
    Input #1, ItemName, Price
    picData.Print ItemName; Tab(25); Format$(Price, "Currency")
     
    For X = 1 To X
        
        
Loop


    
picData.Print

End Sub

