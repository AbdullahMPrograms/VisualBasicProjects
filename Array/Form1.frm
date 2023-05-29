VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picData2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   4080
      ScaleHeight     =   4515
      ScaleWidth      =   3435
      TabIndex        =   5
      Top             =   720
      Width           =   3495
   End
   Begin VB.PictureBox picHeader2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      ScaleHeight     =   555
      ScaleWidth      =   3435
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H8000000E&
      Caption         =   "Read Data"
      Height          =   735
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   4080
      TabIndex        =   2
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox picData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   3435
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.PictureBox picHeader 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdRead_Click()
    'Declarations
    Dim StudentName As String
    Dim StudentAge As Integer
    Dim NumStudents As Integer
    Dim X As Integer
    Dim K As Integer
    
    picHeader.Print "StudentName"; Tab(20); "Age"
    picHeader2.Print "StudentName"; Tab(20); "Age"
    
    'Input/Processing
    Open App.Path & "\INPUTFILE" For Input As #1
    Do While Not EOF(1)
        X = X + 1
        Input #1, StudentName(X), StudentAge(X)
        picData.Print StudentName(X), StudentAge(X)
        Loop
        Close #1
        NumStudents = X
        
        For K = 1 To NumStudents
            If StudentAge(K) >= 18 And StudentAge(K) <= 25 Then
                Print picData2.print; StudentName(K); Tab(20); StudentAge(K)
            End If
        Next K
    
    picNum.Print "Number of Students is: " & Str$(NumStudents)
    
End Sub
