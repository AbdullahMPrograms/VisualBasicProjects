VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "V"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analysis Results"
      Height          =   2175
      Left            =   600
      TabIndex        =   3
      Top             =   1440
      Width           =   3375
      Begin VB.Label lblResult 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Result: "
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblFrownOutput 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Number of Frowns: "
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblSmileOutput 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Smiles: "
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox txtInput 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "&Analyze"
      Height          =   615
      Left            =   4200
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Please Enter a Phrase:"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnalyze_Click()
    'Declarations
    Dim X As Single
    Dim Sentence As String
    Dim PosCount As Integer
    Dim NegCount As Integer
    
    'Input
    Sentence = txtInput.Text
    
    'Processing
    For X = 1 To Len(Sentence)
        If Mid$(Sentence, X, 3) = ":-)" Then
            PosCount = PosCount + 1
        ElseIf Mid$(Sentence, X, 3) = ":-(" Then
            NegCount = NegCount + 1
        End If
    Next X
    
    'Output
    lblSmileOutput.Caption = PosCount
    lblFrownOutput.Caption = NegCount
    
    If PosCount > NegCount Then
        lblResult.Caption = "Happy"
    ElseIf NegCount > PosCount Then
        lblResult.Caption = "Sad"
    ElseIf PosCount = 0 And NegCount = 0 Then
        lblResult.Caption = "None"
    ElseIf PosCount = NegCount Then
        lblResult.Caption = "Unsure"
    End If

End Sub

Private Sub cmdClear_Click()
    txtInput.Text = ""
    lblSmileOutput.Caption = ""
    lblFrownOutput.Caption = ""
    lblResult.Caption = ""
    
End Sub

Private Sub cmdExit_Click()
    End
    
End Sub

