VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Initialize(InputString() As String, ByVal Max As Integer)
    'Declarations
    Dim X As Integer
    
    'Processing
    For X = 1 To Max
        InputString(X) = ""
    Next X
    
End Sub

Public Sub InitializeInteger(Integers() As Integer, ByVal Max As Integer, ByVal Num As Integer)
    'Declarations
    Dim X As Integer
    
    'Processing
    For X = 1 To Max
        Integers(X) = Num
    Next X
    
End Sub

Public Sub Names(InputedNames() As String, ByVal NumNames As Integer, LongestName As String)
    Dim X As Integer
    
    For X = 1 To NumNames
        If Len(InputedNames(X)) > Len(InputedNames(X + 1)) Then
            LongestName = InputedNames(X)
        End If
    Next X
End Sub

Public Sub InitializeDecimals(DecimalNums() As Single, Total As Single, ByVal NumNums As Integer)
    Dim X As Integer
    
    For X = 1 To NumNums
        Total = Total + DecimalNums(X)
    Next X
End Sub

Public Sub Table(Name() As String, Age() As Integer, ByVal NumPpl As Integer)
    Dim X As Integer
    
    For X = 1 To NumPpl
        Print Name(X); Tab(10); Age(X)
    Next X
End Sub

