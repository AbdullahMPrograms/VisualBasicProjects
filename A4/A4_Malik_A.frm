VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test Data Result Analysis"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Display Statistics"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   5
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "Display Results"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Data"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Data File"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   1455
   End
   Begin VB.PictureBox picAnalysis 
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
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   4920
      Width           =   6495
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
      Height          =   4695
      Left            =   240
      ScaleHeight     =   4635
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Const MaxStudents = 50
    
    Dim TestAnswer As String
    Dim StudentName(1 To MaxStudents) As String
    Dim StudentResult(1 To MaxStudents) As String
    Dim RightCount(1 To MaxStudents) As Integer
    Dim WrongCount(1 To MaxStudents) As Integer
    Dim Percentage(1 To MaxStudents) As Single
    Dim Msg As String
    Dim NumStudents As Integer
    Dim HighestScore(1 To MaxStudents) As Integer
    Dim TotalPercentage As Single
    Dim AverageScore As Single
    
Private Sub cmdCalculate_Click()
    'Declarations
    Dim X As Integer
    Dim K As Integer
    
    X = 0
    
    'Processing
    For X = 1 To NumStudents
        For K = 1 To Len(TestAnswer)            'Determine Right Answers and Wrong Answers
            If Mid$(StudentResult(X), K, 1) = Mid$(TestAnswer, K, 1) Then
                RightCount(X) = RightCount(X) + 1
            ElseIf Mid$(StudentResult(X), K, 1) <> Mid$(TestAnswer, K, 1) Then
                WrongCount(X) = WrongCount(X) + 1
            End If
        Next K
    Percentage(X) = RightCount(X) / (RightCount(X) + WrongCount(X))
    TotalPercentage = TotalPercentage + Percentage(X)
    Next X

    AverageScore = TotalPercentage / NumStudents        'Determine Average Score
    
    'Output
    Msg = "Data Has Been Successfully Calculated"
    MsgBox Msg, Vbinformational, "Test Data Results Analysis"
    
    cmdResults.Enabled = True
    
End Sub

Private Sub CmdEnd_Click()
    End
End Sub

Private Sub cmdRead_Click()
    
    'Input
    Open App.Path & "\TestData.txt" For Input As #1
    Input #1, TestAnswer
    Do While Not EOF(1)
        X = X + 1
        Input #1, StudentName(X), StudentResult(X)      'Assign Input to Array
    Loop
    Close #1
    
    NumStudents = X
    Msg = "The File Has Been Successfuly Read"
    
    MsgBox Msg, Vbinformational, "Test Data Result Analysis"
    cmdCalculate.Enabled = True
    
End Sub

Private Sub cmdResults_Click()
    'Setup Header
    picData.Cls
    picData.Print Tab(6); "Student Name"; Tab(25); "Correct"; Tab(34); "Incorrect"; Tab(45); "Percentage"
    
    'Declarations
    Dim X As Integer

    For X = 1 To NumStudents
        picData.Print Tab(2); X & "."; Tab(7); StudentName(X); Tab(26); Format$(RightCount(X), "@@@"); Tab(36); Format$(WrongCount(X), "@@@@"); Tab(48); Format$(Percentage(X), "00.0%")
    Next X
    
    cmdRead.Enabled = False         'Disable Previous Command Buttons to Bypass Error
    cmdCalculate.Enabled = False
    cmdStatistics.Enabled = True
    
End Sub

Private Sub cmdStatistics_Click()
    Dim X As Integer
    Dim HighestScore As Single
    
    picAnalysis.Cls
    HighestScore = 1
    
    For X = 2 To NumStudents
        If RightCount(X) > RightCount(HighestScore) Then
            HighestScore = X            'Determine HighestScore
        End If
    Next X
    picAnalysis.Print "Person with the Highest Score: " & StudentName(HighestScore) & " (" & RightCount(HighestScore) & " out of " & (RightCount(HighestScore) + WrongCount(HighestScore)) & ")"
    picAnalysis.Print "There were " & NumStudents & " students" & " and the average test score was: " & Format$(AverageScore, "00.0%")
End Sub
