' vybe-test: vb/vb_delegates_addressof/delegate_lambda_sub
' origin: languages/vb/tests/vb/test_vb_delegates_addressof.rs

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

Delegate Sub D()
Module M
Sub Main()
Dim del As D = Sub() __Check(CStr("Lambda"), "Lambda")
del()
End Sub
End Module
