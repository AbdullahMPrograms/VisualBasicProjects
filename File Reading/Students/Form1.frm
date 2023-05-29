VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   8400
      Width           =   1815
   End
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
      Left            =   720
      ScaleHeight     =   7395
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   720
      Width           =   4695
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
Dim StudentName As String
Dim HF As String
Dim Mark As Integer

X = 0
picData.Cls
picData.Print "StudentName"; Tab(25); "HF", "Mark"
picData.Print

'Input
Open "F:\ICS\File Reading\Students\Students.txt" For Input As #1
Do While Not EOF(1)
    X = X + 1
    Input #1, StudentName, HF, Mark
    picData.Print StudentName; Tab(25); Format$(HF, "@@@"); Tab(30); Format$(Mark, "@@@")
Loop
Close #1

picData.Print
picData.Print "Total Number of Students: "; X

End Sub
