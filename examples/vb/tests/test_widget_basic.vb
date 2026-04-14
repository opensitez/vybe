' Test 1: Basic controls with positions — no nesting, no colors
' Verifies: control creation, property assignment, Name/Text properties
Dim passed As Integer = 0
Dim failed As Integer = 0

Sub AssertEqual(actual As String, expected As String, testName As String)
    If actual = expected Then
        Console.WriteLine("SUCCESS: " & testName)
        passed = passed + 1
    Else
        Console.WriteLine("FAILURE: " & testName & " — expected '" & expected & "' got '" & actual & "'")
        failed = failed + 1
    End If
End Sub

Public Class BasicForm
    Inherits System.Windows.Forms.Form

    Friend WithEvents lblTitle As Label
    Friend WithEvents btnOk As Button
    Friend WithEvents txtInput As TextBox

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.lblTitle = New Label()
        Me.btnOk = New Button()
        Me.txtInput = New TextBox()

        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Text = "Hello World"
        Me.lblTitle.Location = New Point(10, 10)
        Me.lblTitle.Size = New Size(200, 20)

        Me.btnOk.Name = "btnOk"
        Me.btnOk.Text = "Click Me"
        Me.btnOk.Location = New Point(10, 50)
        Me.btnOk.Size = New Size(100, 30)

        Me.txtInput.Name = "txtInput"
        Me.txtInput.Text = "Type here"
        Me.txtInput.Location = New Point(10, 100)
        Me.txtInput.Size = New Size(200, 25)

        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.txtInput)

        Me.Name = "BasicForm"
        Me.Text = "Test: Basic Controls"
        Me.ClientSize = New Size(400, 300)
    End Sub
End Class

Dim f As New BasicForm()

' Verify control properties were set correctly
AssertEqual(f.lblTitle.Name, "lblTitle", "Label name")
AssertEqual(f.lblTitle.Text, "Hello World", "Label text")
AssertEqual(f.btnOk.Name, "btnOk", "Button name")
AssertEqual(f.btnOk.Text, "Click Me", "Button text")
AssertEqual(f.txtInput.Name, "txtInput", "TextBox name")
AssertEqual(f.txtInput.Text, "Type here", "TextBox text")
AssertEqual(f.Name, "BasicForm", "Form name")
AssertEqual(f.Text, "Test: Basic Controls", "Form text")

Console.WriteLine("")
Console.WriteLine("Test 1 - Basic Controls: " & passed & " passed, " & failed & " failed")
If failed = 0 Then
    Console.WriteLine("SUCCESS: All basic control tests passed")
Else
    Console.WriteLine("FAILURE: Some basic control tests failed")
End If
