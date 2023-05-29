VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   5040
      Width           =   2415
   End
   Begin VB.PictureBox picData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4635
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   120
      Width           =   4815
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
Dim Mark As Integer

'Input
X = 0
picData.Cls

'Processing/Output
Open "F:\ICS\TextMarks" & "\MARKS.TXT" For Input As #1
Do While Not EOF(1)
    X = X + 1
    Input #1, StudentName, Mark
    picData.Print StudentName; Tab(25); Mark
Loop
Close #1
picData.Print "Total Number of Students:    "; X

End Sub
