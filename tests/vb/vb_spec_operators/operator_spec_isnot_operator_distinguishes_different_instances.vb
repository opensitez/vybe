' vybe-test: vb/vb_spec_operators/operator_spec_isnot_operator_distinguishes_different_instances
' origin: languages/vb/tests/vb/test_vb_spec_operators.rs

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

Class Box : End Class
Module M
    Sub Main()
        Dim left As New Box()
        Dim right As New Box()
        __Check(CStr(left IsNot right), "True")
    End Sub
End Module
