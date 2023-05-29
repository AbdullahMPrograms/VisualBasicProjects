VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Q5"
      Height          =   4335
      Left            =   3120
      TabIndex        =   12
      Top             =   4800
      Width           =   2775
      Begin VB.PictureBox picQ5 
         Height          =   3135
         Left            =   240
         ScaleHeight     =   3075
         ScaleWidth      =   2115
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdQ5 
         Caption         =   "Random Words"
         Height          =   615
         Left            =   600
         TabIndex        =   13
         Top             =   3600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Q4"
      Height          =   4335
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   2775
      Begin VB.CommandButton cmdQ4 
         Caption         =   "Random Words"
         Height          =   615
         Left            =   600
         TabIndex        =   11
         Top             =   3600
         Width           =   1455
      End
      Begin VB.PictureBox picQ4 
         Height          =   3135
         Left            =   240
         ScaleHeight     =   3075
         ScaleWidth      =   2115
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Q3"
      Height          =   4335
      Left            =   6120
      TabIndex        =   6
      Top             =   240
      Width           =   2775
      Begin VB.PictureBox picQ3 
         Height          =   3135
         Left            =   240
         ScaleHeight     =   3075
         ScaleWidth      =   2115
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdQ3 
         Caption         =   "Initialize Num()"
         Height          =   615
         Left            =   600
         TabIndex        =   7
         Top             =   3600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Q2"
      Height          =   4335
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   2775
      Begin VB.CommandButton cmdQ2 
         Caption         =   "Print Random Letters"
         Height          =   615
         Left            =   600
         TabIndex        =   5
         Top             =   3600
         Width           =   1455
      End
      Begin VB.PictureBox picQ2 
         Height          =   3135
         Left            =   240
         ScaleHeight     =   3075
         ScaleWidth      =   2115
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdQ1 
      Caption         =   "Print Random Integers"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Q1"
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2775
      Begin VB.PictureBox picQ1 
         Height          =   3135
         Left            =   240
         ScaleHeight     =   3075
         ScaleWidth      =   2115
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQ1_Click()
    Dim X As Integer
    
    For X = 1 To 50
        picQ1.Print Int(Rnd * 100)
    Next X
End Sub

Private Sub cmdQ2_Click()
    Dim Num As Integer
    Dim X As Integer
    
    For X = 1 To 25
        Num = Int(Rnd * (26 - 1 + 1) + 1)
        picQ2.Print Chr$(Num + 64)
    Next X
End Sub

Private Sub cmdQ3_Click()
    Const Max = 1000
    Dim X As Integer
    
    Dim Num(1 To Max) As Integer
    
    For X = 1 To Max
        Num(X) = Int(Rnd * (600 - 100 + 1) + 100)
        picQ3.Print Num(X)
    Next X
    
    
End Sub

Private Sub cmdQ4_Click()
    Const Max = 100 'Const high and low next time
    Dim X As Integer
    Dim Y As Integer
    
    Dim Word(1 To Max) As String
    Dim Char As String
    
    For X = 1 To Max
        Word(X) = ""
        For Y = 1 To Int(Rnd * (15 - 4 + 1) + 4)    'Should be written as variable
            Char = Chr$(Int(Rnd * (26 - 1 + 1) + 1) + 96)
            Word(X) = Word(X) & Char
        Next Y
        picQ4.Print Word(X)
    Next X
        
End Sub

Private Sub cmdQ5_Click()

End Sub

Private Sub Form_Load()
    Randomize
End Sub
