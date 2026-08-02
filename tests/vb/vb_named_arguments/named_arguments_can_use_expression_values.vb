' vybe-test: vb/vb_named_arguments/named_arguments_can_use_expression_values
' origin: languages/vb/tests/vb/test_vb_named_arguments.rs

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

Module M
    Function Compute(total As Integer, scale As Integer, offset As Integer) As Integer
        Return total * scale + offset
    End Function

    Sub Main()
        __Check(CStr(Compute(offset:=3, total:=5 + 1, scale:=2)), "15")
    End Sub
End Module
