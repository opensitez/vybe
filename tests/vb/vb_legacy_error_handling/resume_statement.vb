' vybe-test: vb/vb_legacy_error_handling/resume_statement
' origin: languages/vb/tests/vb/test_vb_legacy_error_handling.rs

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
        Dim attempts As Integer = 0
        
        On Error GoTo Handler
RetryPoint:
        If attempts = 0 Then
            Dim x As Integer = 1 \ 0
        End If
        __Check(CStr("Success"), "Attempt: 1")
        Exit Sub
        
Handler:
        attempts = attempts + 1
        __Check(CStr("Attempt: " & attempts), "Success")
        Resume RetryPoint
    End Sub
End Module
