VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "A5"
   ClientHeight    =   6165
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5880
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
   ScaleHeight     =   6165
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
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
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   5880
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Display"
         Shortcut        =   ^D
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
'Date: March 25 2021
    
    Const Max = 75
    
    Dim StudentName(1 To Max) As String
    Dim StudentAge(1 To Max) As Integer
    Dim X As Integer
    Dim NumStudents As Integer

Private Sub mnuAbout_Click()
    MsgBox "Abdullah Malik, March 25 2021", vbInformation
End Sub

Private Sub mnuAge_Click()
    'Declarations
    Dim Age As Integer
    Dim StudentAgeCount As Integer
    picData.Cls
    
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
    picData.Print Tab(2); "Age: " & Age & " or older."
    picData.Print Tab(2); "Number Selected: " & StudentAgeCount
    
End Sub

Private Sub mnuDisplay_Click()
    'Declarations
    picData.Cls
    
    'Processing/Output
    For X = 1 To NumStudents
        picData.Print Tab(2); StudentName(X); Tab(10); StudentAge(X)
    Next X
    
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuOpen_Click()
    'Declarations
    cdlDialog.FileName = ""
    cdlDialog.InitDir = App.Path
    cdlDialog.Filter = "Text Files|*.txt|All Files|*.*"
    cdlDialog.ShowOpen
    
    'Input
    FileName = cdlDialog.FileName
    
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
    
    'Output
    lblOutput.Caption = "There were " & NumStudents & " records read from " & FileName
End Sub
