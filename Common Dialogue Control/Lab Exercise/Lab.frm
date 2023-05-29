VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Common Dialog"
   ClientHeight    =   8400
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   6840
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picData 
      Height          =   4215
      Left            =   360
      ScaleHeight     =   4155
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   4800
      Width           =   6855
   End
   Begin VB.Menu mnuFule 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Display"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileName As String

Private Sub mnuAbout_Click()
    MsgBox "Menu Design and File Access Lab"
End Sub

Private Sub mnuDisplay_Click()
    'Declarations
    Const Max = 25
    
    Dim CustomerName(1 To Max) As String
    Dim CustomerBirthDate(1 To Max) As String
    Dim CustomerStatus(1 To Max) As Boolean
    Dim X As Integer
    Dim NumCustomer As Integer
    
    'Input
    Open FileName For Input As #1
    Do While Not EOF(1)
        X = X + 1
        Input #1, CustomerName(X), CustomerBirthDate(X), CustomerStatus(X)
        For X = 1 To Max
            picData.Print CustomerName(X), CustomerBirthDate(X), CustomerStatus(X)
            picData.Print ;
        Next X
        NumCustomer = NumCustomer + X
    Loop
    Close #1
        
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
    
    'Output
    MsgBox "File: " & FileName, vbInformation, "User File"
End Sub
