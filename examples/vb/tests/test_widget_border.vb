' Test 4: BorderStyle enum
' Verifies: BorderStyle enum resolves, property is set on panel
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

Sub AssertNotEmpty(actual As String, testName As String)
    If actual <> "" Then
        Console.WriteLine("SUCCESS: " & testName & " (value: " & actual & ")")
        passed = passed + 1
    Else
        Console.WriteLine("FAILURE: " & testName & " — value was empty")
        failed = failed + 1
    End If
End Sub

Public Class BorderTest
    Inherits System.Windows.Forms.Form

    Friend WithEvents pnlBorder As Panel

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.pnlBorder = New Panel()

        Me.pnlBorder.Name = "pnlBorder"
        Me.pnlBorder.Location = New Point(30, 30)
        Me.pnlBorder.Size = New Size(300, 200)
        Me.pnlBorder.BorderStyle = BorderStyle.FixedSingle

        Me.Controls.Add(Me.pnlBorder)

        Me.Name = "BorderTest"
        Me.Text = "Test: BorderStyle"
        Me.ClientSize = New Size(400, 300)
    End Sub
End Class

Dim f As New BorderTest()

' Verify panel properties
AssertEqual(f.pnlBorder.Name, "pnlBorder", "Panel name")

' Verify BorderStyle was assigned (should be 1 = FixedSingle)
Dim bs As String = "" & f.pnlBorder.BorderStyle
AssertEqual(bs, "1", "Panel BorderStyle is FixedSingle (1)")

' Verify form
AssertEqual(f.Name, "BorderTest", "Form name")
AssertEqual(f.Text, "Test: BorderStyle", "Form text")

Console.WriteLine("")
Console.WriteLine("Test 4 - BorderStyle: " & passed & " passed, " & failed & " failed")
If failed = 0 Then
    Console.WriteLine("SUCCESS: All border style tests passed")
Else
    Console.WriteLine("FAILURE: Some border style tests failed")
End If
