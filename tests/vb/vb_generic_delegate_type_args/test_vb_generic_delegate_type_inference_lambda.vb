' vybe-test: vb/vb_generic_delegate_type_args/test_vb_generic_delegate_type_inference_lambda
' origin: languages/vb/tests/vb/test_vb_generic_delegate_type_args.rs

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

Module Program
    Private Function Apply(Of T, R)(val As T, fn As System.Func(Of T, R)) As R
        Return fn(val)
    End Function

    Sub Main()
        Dim res = Apply(5, Function(n) "Scaled_" & (n * 10))
        __Check(CStr(res), "Scaled_50")
    End Sub
End Module
