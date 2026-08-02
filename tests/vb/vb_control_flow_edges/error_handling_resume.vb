' vybe-test: vb/vb_control_flow_edges/error_handling_resume
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

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
        Dim attempts = 0
        On Error GoTo Handler
        
    RetryPoint:
        If attempts = 0 Then
            Throw New System.Exception()
        End If
        __Check(CStr("Success"), "Attempt 1")
        Exit Sub
        
    Handler:
        attempts += 1
        __Check(CStr("Attempt " & attempts), "Success")
        Resume RetryPoint
    End Sub
End Module
