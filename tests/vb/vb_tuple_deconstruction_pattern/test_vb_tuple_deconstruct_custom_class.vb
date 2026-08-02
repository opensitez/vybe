' vybe-test: vb/vb_tuple_deconstruction_pattern/test_vb_tuple_deconstruct_custom_class
' origin: languages/vb/tests/vb/test_vb_tuple_deconstruction_pattern.rs

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

Class Point2D
    Public X As Double
    Public Y As Double

    Public Sub New(x As Double, y As Double)
        Me.X = x
        Me.Y = y
    End Sub

    Public Sub Deconstruct(ByRef outX As Double, ByRef outY As Double)
        outX = Me.X
        outY = Me.Y
    End Sub
End Class

Module Program
    Sub Main()
        Dim pt As New Point2D(3.0, 4.0)
        Dim px, py As Double
        (px, py) = pt
        __Check(CStr(px & "," & py), "3,4")
    End Sub
End Module
