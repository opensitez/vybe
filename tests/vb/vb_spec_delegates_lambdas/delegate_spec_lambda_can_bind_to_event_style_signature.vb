' vybe-test: vb/vb_spec_delegates_lambdas/delegate_spec_lambda_can_bind_to_event_style_signature
' origin: languages/vb/tests/vb/test_vb_spec_delegates_lambdas.rs

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
        Dim handler As Action(Of Object, EventArgs) = Sub(sender As Object, e As EventArgs)
            __Check(CStr("handled"), "handled")
        End Sub
        handler(Nothing, Nothing)
    End Sub
End Module
