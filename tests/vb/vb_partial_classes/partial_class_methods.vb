' vybe-test: vb/vb_partial_classes/partial_class_methods
' origin: languages/vb/tests/vb/test_vb_partial_classes.rs

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

Partial Class Calculator
    Public Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
End Class

Partial Class Calculator
    Public Function Multiply(a As Integer, b As Integer) As Integer
        Return a * b
    End Function
End Class

Module M
    Sub Main()
        Dim c As New Calculator()
        __Check(CStr(c.Add(2, 3)), "5")
        __Check(CStr(c.Multiply(2, 3)), "6")
    End Sub
End Module
