VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Form1"
   ClientHeight    =   280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   340
   LinkTopic       =   "Form1"
   ScaleHeight     =   280
   ScaleWidth      =   340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGrade 
      Text            =   "85"
      Height          =   25
      Left            =   50
      TabIndex        =   0
      Top             =   100
      Width           =   150
      BackColor       =   &H00FCFAF8&
      ForeColor       =   &H002A170F&
      Font            =   "Segoe UI, 12px"
      Enabled         =   -1
      Visible         =   -1
   End
   Begin VB.CommandButton btnGrade 
      Caption         =   "Calc Grade"
      Height          =   30
      Left            =   50
      TabIndex        =   0
      Top             =   150
      Width           =   140
      BackColor       =   &H00FCFAF8&
      ForeColor       =   &H002A170F&
      Font            =   "Segoe UI, 12px"
      Enabled         =   -1
      Visible         =   -1
   End
   Begin VB.CommandButton btnSum 
      Caption         =   "Sum Grades"
      Height          =   30
      Left            =   50
      TabIndex        =   0
      Top             =   210
      Width           =   150
      BackColor       =   &H00FCFAF8&
      ForeColor       =   &H002A170F&
      Font            =   "Segoe UI, 12px"
      Enabled         =   -1
      Visible         =   -1
   End
   Begin VB.Label lbl1 
      Caption         =   "Shows use of select (Calc) and arrays (Sum)"
      Height          =   20
      Left            =   20
      TabIndex        =   0
      Top             =   60
      Width           =   280
      BackColor       =   &H00FCFAF8&
      ForeColor       =   &H002A170F&
      Font            =   "Segoe UI, 12px"
      Enabled         =   -1
      Visible         =   -1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnGrade_Click()

    ' Test Select Case
    Dim grade As Integer
    grade = val(txtGrade.text)
    Select Case grade
        Case 90 To 100
            MsgBox "Grade: A - Excellent!"
        Case 80 To 89
            MsgBox "Grade: B - Good work!"
        Case 70 To 79
            MsgBox "Grade: C"
        Case 60 To 69
            MsgBox "Grade: D"
        Case Else
            MsgBox "Grade: F"
    End Select
End Sub

Private Sub btnSum_Click()
    Dim numbers() As Integer = {10, 20, 30, 40, 50}
    ' Test Arrays
    ' MsgBox "Array element 2: " & CStr(numbers(2))
    ' MsgBox "Array size: " & CStr(UBound(numbers) + 1)

    ' Test array with loop
    Dim i As Integer
    Dim total As Integer
    total = 0

    For i = 0 To UBound(numbers)
        total = total + numbers(i)
    Next i

    MsgBox "Sum of all numbers: " & CStr(total)
End Sub