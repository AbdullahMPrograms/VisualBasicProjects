VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "A6"
   ClientHeight    =   5835
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   7320
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lblOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   5655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAge 
         Caption         =   "&Age Selection"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer Name: Abdullah Malik
'Program Purpose: Read and Display Student Names and Ages from a text file, Display ages above user input only when age is entered.
'Date: April 7 2021
    
    Const Max = 75
    
    Dim StudentName(1 To Max) As String
    Dim StudentAge(1 To Max) As Integer
    Dim X As Integer
    Dim NumStudents As Integer
    Dim FileName As String

Private Sub mnuAbout_Click()
    MsgBox "Abdullah Malik, March 25 2021", vbInformation
End Sub

Private Sub mnuAge_Click()
    'Declarations
    Dim Age As Integer
    Dim StudentAgeCount As Integer
    picData.Cls
    
    X = 0
    StudentAgeCount = 0
    
    'Input
    Age = Val(InputBox$("Please input an Age"))     'Determine Age Value
    
    'Processing
    For X = 1 To NumStudents
        If StudentAge(X) >= Age Then
            picData.Print Tab(2); StudentName(X); Tab(10); StudentAge(X)    'Print Age and Name of student older than inputed value
            StudentAgeCount = StudentAgeCount + 1
        End If
    Next X
    
    'Output
    picData.Print
    picData.Print Tab(2); "Age: " & Str$(Age) & " or older."
    picData.Print Tab(2); "Number Selected: " & Str$(StudentAgeCount)
    
End Sub


Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuOpen_Click()
    
    GetFile FileName
    ReadData FileName, NumStudents, StudentName(), StudentAge()
    DisplayData NumStudents, StudentName(), StudentAge()
    
End Sub

Public Sub GetFile(FName As String)
    'Declarations
    cdlDialog.FileName = ""
    cdlDialog.InitDir = App.Path
    cdlDialog.Filter = "Text Files|*.txt|All Files|*.*"
    cdlDialog.ShowOpen
    
    'Input
    FileName = cdlDialog.FileName
    
End Sub

Public Sub ReadData(ByVal FileName As String, Num As Integer, StudentN() As String, StudentA() As Integer)
    'Declarations
    X = 0
    NumStudents = 0
    
    'Processing
    If FileName = "" Then
        MsgBox "No File was Selected", vbCritical       'Bypass error when no file is selected
    Else
        Open FileName For Input As #1
        Do While Not EOF(1)
            X = X + 1
            Input #1, StudentName(X), StudentAge(X)     'Assign values
        Loop
        Close #1
    End If
    
    NumStudents = NumStudents + X
End Sub

Public Sub DisplayData(ByVal Num As Integer, StudentN() As String, StudentA() As Integer)
    'Declarations
    picData.Cls
    X = 0
    
    'Processing/Output
    For X = 1 To NumStudents
        picData.Print Tab(2); StudentName(X); Tab(10); StudentAge(X)
    Next X
    
    'Output
    lblOutput.Caption = "There were " & Str$(NumStudents) & " records read from " & FileName
End Sub

