' Test 2: Nested container — panel with child controls
' Verifies: panel creation, child controls added to panel, property assignment
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

Public Class NestedTest
    Inherits System.Windows.Forms.Form

    Friend WithEvents pnlRoot As Panel
    Friend WithEvents lblInside As Label
    Friend WithEvents btnInside As Button

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.pnlRoot = New Panel()
        Me.lblInside = New Label()
        Me.btnInside = New Button()

        Me.pnlRoot.Name = "pnlRoot"
        Me.pnlRoot.Location = New Point(50, 50)
        Me.pnlRoot.Size = New Size(300, 200)

        Me.lblInside.Name = "lblInside"
        Me.lblInside.Text = "I am inside the panel"
        Me.lblInside.Location = New Point(10, 10)
        Me.lblInside.Size = New Size(200, 20)

        Me.btnInside.Name = "btnInside"
        Me.btnInside.Text = "Panel Button"
        Me.btnInside.Location = New Point(10, 50)
        Me.btnInside.Size = New Size(120, 30)

        ' NESTING: add children to panel, not form
        Me.pnlRoot.Controls.Add(Me.lblInside)
        Me.pnlRoot.Controls.Add(Me.btnInside)
        Me.Controls.Add(Me.pnlRoot)

        Me.Name = "NestedTest"
        Me.Text = "Test: Nested Container"
        Me.ClientSize = New Size(500, 400)
    End Sub
End Class

Dim f As New NestedTest()

' Verify panel properties
AssertEqual(f.pnlRoot.Name, "pnlRoot", "Panel name")

' Verify child controls exist and have correct properties
AssertEqual(f.lblInside.Name, "lblInside", "Nested label name")
AssertEqual(f.lblInside.Text, "I am inside the panel", "Nested label text")
AssertEqual(f.btnInside.Name, "btnInside", "Nested button name")
AssertEqual(f.btnInside.Text, "Panel Button", "Nested button text")

' Verify form
AssertEqual(f.Name, "NestedTest", "Form name")
AssertEqual(f.Text, "Test: Nested Container", "Form text")

Console.WriteLine("")
Console.WriteLine("Test 2 - Nested Container: " & passed & " passed, " & failed & " failed")
If failed = 0 Then
    Console.WriteLine("SUCCESS: All nested container tests passed")
Else
    Console.WriteLine("FAILURE: Some nested container tests failed")
End If
