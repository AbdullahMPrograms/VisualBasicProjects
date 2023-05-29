VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   13485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   ScaleHeight     =   13485
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Caption         =   "Q9: Perfect Squares"
      Height          =   2535
      Left            =   9960
      TabIndex        =   41
      Top             =   6000
      Width           =   4215
      Begin VB.CommandButton cmdPerfectSquare 
         Caption         =   "Calculate Perfect Squares"
         Height          =   1215
         Left            =   360
         TabIndex        =   42
         Top             =   720
         Width           =   3615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Q5: Federal Political Tax Credit"
      Height          =   2535
      Left            =   600
      TabIndex        =   35
      Top             =   6000
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   420
         Left            =   1800
         TabIndex        =   37
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Calculate Credit"
         Height          =   615
         Left            =   360
         TabIndex        =   36
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "Contribution"
         Height          =   375
         Left            =   480
         TabIndex        =   38
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Q7: How Many Wolverines?"
      Height          =   5175
      Left            =   5280
      TabIndex        =   32
      Top             =   6000
      Width           =   4215
      Begin VB.CommandButton cmdCls 
         Caption         =   "Clear Screen"
         Height          =   495
         Left            =   960
         TabIndex        =   40
         Top             =   4560
         Width           =   2295
      End
      Begin VB.PictureBox picOutput 
         BackColor       =   &H8000000E&
         Height          =   3015
         Left            =   360
         ScaleHeight     =   2955
         ScaleWidth      =   3555
         TabIndex        =   39
         Top             =   1440
         Width           =   3615
      End
      Begin VB.CommandButton cmdWolverine 
         Caption         =   "How Many Wolverines?"
         CausesValidation=   0   'False
         Height          =   615
         Left            =   360
         TabIndex        =   33
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label19 
         Height          =   375
         Left            =   480
         TabIndex        =   34
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Q6: Whats the Greatest?"
      Height          =   2535
      Left            =   600
      TabIndex        =   29
      Top             =   8640
      Width           =   4215
      Begin VB.CommandButton cmdGreatest 
         Caption         =   "Whats the Greatest?"
         CausesValidation=   0   'False
         Height          =   1095
         Left            =   480
         TabIndex        =   30
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label18 
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Q4: Price of Carpet"
      Height          =   5175
      Left            =   9960
      TabIndex        =   18
      Top             =   480
      Width           =   4215
      Begin VB.CommandButton cmdClearAll2 
         Caption         =   "C&lear All"
         Height          =   615
         Left            =   480
         TabIndex        =   28
         Top             =   4320
         Width           =   3375
      End
      Begin VB.TextBox txtQuantity 
         Height          =   420
         Left            =   1800
         TabIndex        =   20
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmdCalculatePrice 
         Caption         =   "Calculate &Price"
         Height          =   615
         Left            =   360
         TabIndex        =   19
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label17 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   480
         TabIndex        =   27
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Sub-total:"
         Height          =   375
         Left            =   600
         TabIndex        =   26
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "HST:"
         Height          =   495
         Left            =   600
         TabIndex        =   25
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Total:"
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblSubTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   23
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblHST 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblTotal2 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   3600
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   5520
      TabIndex        =   17
      Top             =   11280
      Width           =   3735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Q3: Whats the Average?"
      Height          =   2535
      Left            =   5280
      TabIndex        =   14
      Top             =   3120
      Width           =   4215
      Begin VB.CommandButton cmdAverage 
         Caption         =   "Whats the &Average?"
         Height          =   1095
         Left            =   480
         TabIndex        =   15
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Q2: Can You Drive or Drink?"
      Height          =   2535
      Left            =   5280
      TabIndex        =   11
      Top             =   480
      Width           =   4215
      Begin VB.CommandButton cmdAge 
         Caption         =   "&Can You Drive or Drink?"
         Height          =   1095
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Q1: Convert to Seconds"
      Height          =   5175
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   4215
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "C&lear All"
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox txtSecond 
         Height          =   420
         Left            =   1800
         TabIndex        =   9
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtHour 
         Height          =   465
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtMinute 
         Height          =   420
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdCalculateSeconds 
         Caption         =   "Calculate &Seconds"
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   2760
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Seconds:"
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Hours:"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Minutes:"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Total Seconds:"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   3840
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAge_Click()
    'Declarations
    Dim Age As Single
    Dim Msg As String
    
    'Input
    Age = Val(InputBox$("Please Enter Your Age", "Are you 18?"))
    
    'Processing
    If Age >= 19 Then
        Msg = "You are old enough to drive and drink!"
    ElseIf Age <= 15 Then
        Msg = "You can’t drive or drink!"
    Else
        Msg = "You are old enough to drive but not drink!"
    End If
    
    'Output
    MsgBox Msg, vbInformation, "Are you 18?"
    
End Sub

Private Sub cmdAverage_Click()
    'Declarations
    Dim X As Integer
    Dim Num As Integer
    Dim Sum As Integer
    Dim Numbers As Integer
    Dim Average As Single
    
    'Input
    Numbers = Val(InputBox$("Please Enter the Number of Numbers", "Whats the Average"))
    
    'Processing
    Sum = 0
    For X = 1 To Numbers
        Num = Val(InputBox$("Please Enter Numbers" & Str$(X) & ":"))
        Sum = Sum + Num
    Next X
    If Numbers > 0 Then
        Average = Sum / Numbers
        Msg = "Average of Given Numbers is: " & Format$(Average, "###,###,##0.0")
    Else
        Msg = "No Numbers were Given"
    End If
    
    'Output
    MsgBox Msg, vbInformation, "Whats the Average?"
    
End Sub

Private Sub cmdCalculatePrice_Click()
    'Declarations
    Dim Price As Single
    Dim Quantity As Single
    Dim SubTotal As Single
    Dim HST As Single
    Dim Total As Single
    
    'Input
    Quantity = Val(txtQuantity.Text)
    If Quantity >= 24 Then
        Price = 18
    ElseIf Quantity <= 8 Then
        Price = 25
    Else
        Price = 21
    End If
    
    'Processing
    SubTotal = Price * Quantity + 75
    HST = SubTotal * 0.13
    Total = SubTotal + HST
    
    'Output
    lblSubTotal.Caption = FormatCurrency(SubTotal)
    lblHST.Caption = FormatCurrency(HST)
    lblTotal2.Caption = FormatCurrency(Total)

End Sub

Private Sub cmdCalculateSeconds_Click()
    'Declarations
    Dim HourinSecond As Single
    Dim MinuteinSecond As Single
    Dim TotalSecond As Single
    Dim Second As Single
    Dim Total As Double
    
    'Input
    HourinSecond = Val(txtHour.Text) * 3600
    MinuteinSecond = Val(txtMinute.Text) * 60
    Second = Val(txtSecond.Text)
    
    'Processing
    TotalSecond = HourinSecond + MinuteinSecond + Second
    
    'Output
    lblTotal.Caption = Format$(TotalSecond, "###,###,###,##0")
    
End Sub

Private Sub cmdClearAll_Click()
        txtHour.Text = ""
        txtMinute.Text = ""
        txtSecond.Text = ""
        lblTotal.Caption = ""
End Sub

Private Sub cmdClearAll2_Click()
        txtQuantity.Text = ""
        lblHST.Caption = ""
        lblSubTotal.Caption = ""
        lblTotal2.Caption = ""
End Sub

Private Sub cmdCls_Click()
    picOutput.Cls
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdGreatest_Click()
    'Declarations
    Dim Num1 As Single
    Dim Num2 As Single
    Dim Num3 As Single
    Dim Largest As Single
    
    'Input
    Num1 = Val(InputBox$("Input Number 1", "Whats the Greatest?"))
    Num2 = Val(InputBox$("Input Number 2", "Whats the Greatest?"))
    Num3 = Val(InputBox$("Input Number 3", "Whats the Greatest?"))
    Largest = 0
    
    'Processing
    If Num1 > Largest Then
        Largest = Num1
    End If
    If Num2 > Largest Then
        Largest = Num2
    End If
    If Num3 > Largest Then
        Largest = Num3
    End If
    
    'Output
    Msg = "The Greatest Number is: " & Format$(Largest, "###,##0.0")
    MsgBox Msg, vbInformational, "Whats the Greatest?"
    
End Sub

Private Sub cmdPerfectSquare_Click()
    'Declarations
    Dim Number As Single
    Dim X As Integer
    Dim Msg As String
    Dim Output As String
    
    'Processing
    Number = Val(InputBox$("Enter Number", "Perfect Squares"))
    For X = 1 To Int(Sqr(Number))
        Output = Output & Str$(X ^ 2)
    Next X
    
    'Output
    Msg = "The Perfect Squares were: " & Output
    MsgBox Msg, vbInformational, "Perfect Square"
    
End Sub

Private Sub cmdWolverine_Click()
    'Declarations
    Dim Count As Integer
    Dim X As Integer
    
    'Input
    Count = Val(InputBox$("Input Amount of Wolverines", "How Many Wolverines?"))
    
    'Processing
    For X = 1 To Count
        picOutput.Print "Wolverine"
    Next X
    
End Sub

