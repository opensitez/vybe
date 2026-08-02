' vybe-test: vb/vb_inheritance/t08_three_level_inheritance
' origin: languages/vb/tests/vb/vb_inheritance_test.rs

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

Class GrandParent
    Public A As String = "GP"
End Class

Class Parent
    Inherits GrandParent
    Public B As String = "P"
End Class

Class Child
    Inherits Parent
    Public C As String = "C"
End Class

Dim c As New Child()
__Check(CStr(c.A), "GP")
__Check(CStr(c.B), "P")
__Check(CStr(c.C), "C")
