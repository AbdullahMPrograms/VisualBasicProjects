VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQ6 
      Caption         =   "Q6"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdQ5 
      Caption         =   "Q5"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdQ4 
      Caption         =   "Q4"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdQ3 
      Caption         =   "Q3"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdQ2 
      Caption         =   "Q2"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdQ1 
      Caption         =   "Q1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox picData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Output(Ch As String)
    If (Ch > "a" And Ch < "z") Or (Ch > "A" And Ch < "Z") Then
        Ch = "."
    ElseIf Ch > "0" And Ch < "9" Then
        Ch = "-"
    Else
        Ch = "?"
    End If
    
End Sub

Public Sub Roundto(Y As Integer, Formatt As String)
    Dim Count As String
    
    For K = 1 To Y
        Count = Count + "0"
        Formatt = "0." & Count
    Next K
    
End Sub

Private Sub cmdQ1_Click()
    'Declarations
    Dim X As Single
    Dim Y As Integer
    Dim Formatt As String
    
    X = InputBox$("Enter a Real Number")
    Y = InputBox$("Enter a Positive Integer")
    
    Roundto Y, Formatt
    
    picData.Print Format$(X, Formatt)
    
End Sub

Private Sub cmdQ2_Click()
    'Declarations
    Dim Ch As String
    
    Ch = InputBox$("Please enter a Single Character")
    
    Output Ch
    
    picData.Print Ch
    
End Sub

Private Sub cmdQ3_Click()
    'Declarations
    Dim N As Integer
    
    N = InputBox$("Enter a Positive Integer")
    
    Sum N
    
    picData.Print N
    
End Sub

Public Sub Sum(N As Integer)
    Dim K As Integer
    Dim Sum As Integer
    
    For K = 1 To N
        Sum = Sum + K
    Next K
    
    N = Sum
End Sub

Private Sub cmdQ4_Click()
    'Declarations
    Dim Sentence As String
    
    Sentence = InputBox$("Enter a Sentence")
    
    Output2 Sentence
    
    picData.Print Sentence
    
End Sub

Public Sub Output2(Sentence As String)
    Dim X As Integer
    Dim NewPhrase As String
    
    NewPhrase = ""
    
    For X = 1 To Len(Sentence)
        Ch = Mid$(Sentence, X, 1)
        If (Ch > "a" And Ch < "z") Or (Ch > "A" And Ch < "Z") Then
            Ch = "."
        ElseIf Ch > "0" And Ch < "9" Then
            Ch = "-"
    End If
        
    NewPhrase = NewPhrase + Ch
    Next X
    
    Sentence = NewPhrase
End Sub


Private Sub cmdQ5_Click()
    'Declarations
    Dim InputDate As String
    
    Dim Day As String
    Dim Month As String
    Dim Year As String
    
    InputDate = InputBox$("Enter a Date in Format: dd/mm/yy")
    
    Day = Mid$(InputDate, 1, 2)
    Month = Mid$(InputDate, 4, 2)
    Year = Mid$(InputDate, 7, 2)
    
    Convert Month, Year
    
    picData.Print Day; " "; Month; " "; Year
End Sub


Public Sub Convert(Month As String, Year As String)
    
    If Month = "01" Then
        Month = "JAN"
    ElseIf Month = "02" Then
        Month = "FEB"
    ElseIf Month = "03" Then
        Month = "MAR"
    ElseIf Month = "04" Then
        Month = "APR"
    ElseIf Month = "05" Then
        Month = "MAY"
    ElseIf Month = "06" Then
        Month = "JUNE"
    ElseIf Month = "07" Then
        Month = "JULY"
    ElseIf Month = "08" Then
        Month = "AUG"
    ElseIf Month = "09" Then
        Month = "SEP"
    ElseIf Month = "10" Then
        Month = "OCT"
    ElseIf Month = "11" Then
        Month = "NOV"
    ElseIf Month = "12" Then
        Month = "DEC"
    End If
    
    Year = "20" & Year
End Sub

Private Sub cmdQ6_Click()
'A = 7, B = 10
'X Remains the same, Y = Y + X

End Sub
