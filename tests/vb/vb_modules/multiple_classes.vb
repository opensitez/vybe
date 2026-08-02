' vybe-test: vb/vb_modules/multiple_classes
' origin: languages/vb/tests/vb/test_vb_modules.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Class Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
End Class

Class Segment
    Public Start As Point
    Public Finish As Point
    Public Sub New(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
        Start = New Point(x1, y1)
        Finish = New Point(x2, y2)
    End Sub
    Public Function Length() As Double
        Dim dx As Integer = Finish.X - Start.X
        Dim dy As Integer = Finish.Y - Start.Y
        Return Math.Sqrt(dx * dx + dy * dy)
    End Function
End Class

Module M
    Sub Main()
        Dim l As New Segment(0, 0, 3, 4)
        __Check(CStr(l.Length()), "5")
    End Sub
End Module
