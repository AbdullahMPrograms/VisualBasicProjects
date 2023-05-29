VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   7080
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDisplay 
      Caption         =   "Display"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About.."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuAbout_Click()
    MsgBox "Menu Design and File Access"
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuOpen_Click()
    Dim FileName As String
    
    cdlDialog.FileName = ""
    cdlDialog.InitDir = App.Path
    cdlDialog.Filter = "Text Files|*.txt|All Files|*.*"
    cdlDialog.ShowOpen
    
    FileName = cdlDialog.FileName
    
    MsgBox "File: " & FileName, vbInformation, "User File"

End Sub
