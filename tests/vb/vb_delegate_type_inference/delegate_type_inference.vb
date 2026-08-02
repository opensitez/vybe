' vybe-test: vb/vb_delegate_type_inference/delegate_type_inference
' origin: languages/vb/tests/vb/test_vb_delegate_type_inference.rs

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
    Sub DoWork(action As Action)
        action()
    End Sub

    Sub Main()
        ' Delegate inference from Lambda
        DoWork(Sub() __Check(CStr("Inferred"), "Inferred"))
    End Sub
End Module
