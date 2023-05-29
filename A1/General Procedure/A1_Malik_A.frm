VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Riverdale Airlines"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   7320
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Height          =   6735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "&Clear All"
         Height          =   495
         Left            =   2760
         TabIndex        =   10
         Top             =   6000
         Width           =   3015
      End
      Begin VB.Frame Frame3 
         Caption         =   "Additional Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   4680
         TabIndex        =   9
         Top             =   840
         Width           =   3135
         Begin VB.CheckBox chkOnboard 
            Caption         =   "Onboard Meal - $10"
            Height          =   615
            Left            =   240
            TabIndex        =   17
            Top             =   1920
            Width           =   2415
         End
         Begin VB.CheckBox chkSeat 
            Caption         =   "Advanced Seat Selection - $20"
            Height          =   615
            Left            =   240
            TabIndex        =   16
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkLuggage 
            Caption         =   "Luggage - $50"
            Height          =   615
            Left            =   240
            TabIndex        =   15
            Top             =   480
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ticket Destination"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   600
         TabIndex        =   7
         Top             =   840
         Width           =   3015
         Begin VB.OptionButton optAsia 
            Caption         =   "Asia/South America - $2100"
            Height          =   615
            Left            =   240
            TabIndex        =   14
            Top             =   1920
            Width           =   2415
         End
         Begin VB.OptionButton optEurope 
            Caption         =   "Europe - $1200"
            Height          =   615
            Left            =   240
            TabIndex        =   13
            Top             =   1200
            Width           =   2415
         End
         Begin VB.OptionButton optNorthAmerica 
            Caption         =   "North America - $500"
            Height          =   615
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   2415
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Riverdale Airlines"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   8
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label lblSubTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Sub-total:"
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "HST:"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Total:"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label lblHST 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   5400
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim TicketPrice As Single
    Dim ExtrasPrice As Single
    Dim SubTotal As Single
    Dim HST As Single
    Dim Total As Single
Public Sub CalculateOptions()
    ExtrasPrice = 0
    If chkLuggage.Value = 1 Then            'ExtrasPrice
        ExtrasPrice = ExtrasPrice + 50
    End If
    
    If chkSeat.Value = 1 Then
        ExtrasPrice = ExtrasPrice + 20
    End If
    
    If chkOnboard.Value = 1 Then
        ExtrasPrice = ExtrasPrice + 10
    End If
    
End Sub

Public Sub CalculateTicket()
    'Input
    TicketPrice = 0
    If optNorthAmerica.Value = True Then    'TicketPrice
        TicketPrice = 500
    ElseIf optEurope.Value = True Then
        TicketPrice = 1200
    ElseIf optAsia.Value = True Then
        TicketPrice = 2100
    End If
    
End Sub

Public Sub Display()
    'Processing
    SubTotal = TicketPrice + ExtrasPrice
    HST = SubTotal * 0.045
    Total = SubTotal + HST
    
    'Output
    lblSubTotal.Caption = Format$(SubTotal, "Currency")
    lblHST.Caption = Format$(HST, "Currency")
    lblTotal.Caption = Format$(Total, "Currency")
    
End Sub

Private Sub chkLuggage_Click()
    CalculateOptions
    Display
End Sub

Private Sub chkOnboard_Click()
    CalculateOptions
    Display
End Sub

Private Sub chkSeat_Click()
    CalculateOptions
    Display
End Sub

Private Sub cmdClearAll_Click()
    optNorthAmerica.Value = False   'Clear Options
    optEurope.Value = False
    optAsia.Value = False
    
    chkLuggage.Value = 0            'Clear Checkboxes
    chkSeat.Value = 0
    chkOnboard.Value = 0
    
    TicketPrice = 0
    ExtrasPrice = 0
    
    lblSubTotal.Caption = ""        'Clear Labels
    lblHST.Caption = ""
    lblTotal.Caption = ""
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub optAsia_Click()
    CalculateTicket
    Display
End Sub

Private Sub optEurope_Click()
    CalculateTicket
    Display
End Sub

Private Sub optNorthAmerica_Click()
    CalculateTicket
    Display
End Sub
