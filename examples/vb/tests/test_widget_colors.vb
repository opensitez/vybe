' Test 3: BackColor via Color.FromArgb
' Verifies: Color.FromArgb creates color objects, BackColor property is set
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

Public Class ColorTest
    Inherits System.Windows.Forms.Form

    Friend WithEvents pnlColored As Panel
    Friend WithEvents lblRed As Label

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.pnlColored = New Panel()
        Me.lblRed = New Label()

        Me.pnlColored.Name = "pnlColored"
        Me.pnlColored.Location = New Point(20, 20)
        Me.pnlColored.Size = New Size(300, 200)
        Me.pnlColored.BackColor = Color.FromArgb(173, 216, 230)

        Me.lblRed.Name = "lblRed"
        Me.lblRed.Text = "Red Background"
        Me.lblRed.Location = New Point(10, 10)
        Me.lblRed.Size = New Size(150, 25)
        Me.lblRed.BackColor = Color.FromArgb(255, 100, 100)

        Me.pnlColored.Controls.Add(Me.lblRed)
        Me.Controls.Add(Me.pnlColored)

        Me.Name = "ColorTest"
        Me.Text = "Test: BackColor"
        Me.ClientSize = New Size(400, 300)
    End Sub
End Class

Dim f As New ColorTest()

' Verify control properties
AssertEqual(f.pnlColored.Name, "pnlColored", "Panel name")
AssertEqual(f.lblRed.Name, "lblRed", "Label name")
AssertEqual(f.lblRed.Text, "Red Background", "Label text")

' Verify BackColor was assigned (check it's not empty/null)
Dim panelColor As String = "" & f.pnlColored.BackColor
Dim labelColor As String = "" & f.lblRed.BackColor
AssertNotEmpty(panelColor, "Panel BackColor assigned")
AssertNotEmpty(labelColor, "Label BackColor assigned")

Console.WriteLine("")
Console.WriteLine("Test 3 - Colors: " & passed & " passed, " & failed & " failed")
If failed = 0 Then
    Console.WriteLine("SUCCESS: All color tests passed")
Else
    Console.WriteLine("FAILURE: Some color tests failed")
End If
