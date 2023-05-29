VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Toys, Toys, Toys"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9285
   Icon            =   "A3_Malik_A.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
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
      ScaleWidth      =   8715
      TabIndex        =   3
      Top             =   120
      Width           =   8775
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
      ScaleWidth      =   8715
      TabIndex        =   2
      Top             =   720
      Width           =   8775
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H8000000E&
      Caption         =   "Read Data"
      Height          =   735
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Total Sales:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label lblTotalPrice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer Name: Abdullah Malik
'Date: March 9th, 2021
'Purpose: Read and display contents of a text file as well as calculating average and number of lines.

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdRead_Click()
'Declarations
'Dim FileDir As String
'Dim FileName As String
Dim ToyName As String
Dim ToyPrice As Single
Dim TotalToyPrice As Single
Dim ToyCount As Integer
Dim RetailSales As Integer
Dim OnlineSales As Integer
Dim ToysSold As Integer
Dim TotalToysSold As Integer
Dim TotalSales As Single
Dim AvgPrice As Single

TotalSales = 0
ToyCount = 0
TotalToysSold = 0

picData.Cls
picHeader.Cls
picHeader.Print Tab(2); "TOY"; Tab(18); "TOY", "RETAIL", "ONLINE", "TOTAL", "TOTAL TOY"     'Header Input
picHeader.Print Tab(2); "NAME"; Tab(18); "PRICE", "SALES", "SALES", "SOLD", "SALES"         '^

'Input/Processsing
'FileDir = InputBox$("Please Enter Path to Text File", "Toys, Toys, Toys")   'Determine File Dir\ and Name
'FileName = InputBox$("Please Enter File Name As .txt", "Toys, Toys, Toys")
'If FileDir = "" Then
    'MsgBox "No Path Was Entered", vbCritical, "Toys, Toys, Toys"
'End If

'If FileName = "" Then
    'MsgBox "No File Name Was Entered", vbCritical, "Toys, Toys, Toys"
'End If
    
'Open (FileDir & "\" & FileName) For Input As #1
Open App.Path & "\Toys.txt" For Input As #1         'Dumb way to get path
Do While Not EOF(1)
    ToyCount = ToyCount + 1                                 'Process all data
    Input #1, ToyName, ToyPrice, RetailSales, OnlineSales
    
    ToysSold = RetailSales + OnlineSales
    
    If OnlineSales <= 50 Then
        TotalToyPrice = ToysSold * ToyPrice
    ElseIf OnlineSales > 50 Then
        TotalToyPrice = (50 + ((ToysSold - 50) * 0.9)) * ToyPrice     '10% Discount
    End If
    
    TotalToysSold = TotalToysSold + ToysSold
    
    picData.Print Tab(2); ToyName; Tab(18); Format$(Format$(ToyPrice, "$0.00"), "@@@@@@"), Format$(RetailSales, "@@@@"), Format$(OnlineSales, "@@@@"), Format$(ToysSold, "@@@@"), Format$(Format$(TotalToyPrice, "$0.00"), "@@@@@@@@")
    
    TotalSales = TotalSales + TotalToyPrice
Loop
Close #1

AvgPrice = TotalSales / TotalToysSold

'Output
lblTotalPrice.Caption = Format$(Format$(TotalSales, "$#,##0.00"), "@@@@@@")

picData.Print
picData.Print Tab(2); "Number of Toys For Sale is: " & ToyCount
picData.Print Tab(2); "Average Price of Toy is " & Format$(Format$(AvgPrice, "$0.00"), "@@@@@")


End Sub

