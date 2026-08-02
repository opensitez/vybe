' vybe-test: vb/vb_spec_delegates_lambdas/delegate_spec_lambda_can_return_object_reference
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

Class Box
    Public Value As String = "box"
End Class
Module M
    Sub Main()
        Dim fn As Func(Of Box) = Function() New Box()
        __Check(CStr(fn().Value), "box")
    End Sub
End Module
