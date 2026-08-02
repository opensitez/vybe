' vybe-test: vb/vb_lambda_byval_args/lambda_byval_args
' origin: languages/vb/tests/vb/test_vb_lambda_byval_args.rs

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
    Sub Main()
        ' Explicit ByVal in lambda arguments
        Dim act As Action(Of Integer) = Sub(ByVal x As Integer)
                                            x += 10
                                        End Sub
        Dim val = 5
        act(val)
        __Check(CStr(val), "5") ' Should still be 5
    End Sub
End Module
