' vybe-test: vb/vb_select_case_statements/select_case_nested
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

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
Dim x = 1, y = 2
Select Case x
Case 1
Select Case y
Case 2
__Check(CStr("Nested"), "Nested")
End Select
End Select
End Sub
End Module
