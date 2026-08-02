' vybe-test: vb/vb_inheritance/t12_no_explicit_ctor_parent_auto_called
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

Class Base
    Public Ready As String = "yes"

    Sub New()
        __Check(CStr("base ctor"), "base ctor")
    End Sub
End Class

Class Child
    Inherits Base
End Class

Dim c As New Child()
__Check(CStr(c.Ready), "yes")
