' vybe-test: vb/vb_multiple_statements_colon/multiple_statements_colon
' origin: languages/vb/tests/vb/test_vb_multiple_statements_colon.rs

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
        ' The colon character allows placing multiple statements on a single line
        Dim x As Integer = 10 : Dim y As Integer = 20 : __Check(CStr(x + y), "30")
        
        If x = 10 Then : __Check(CStr("Yes"), "Yes") : End If
    End Sub
End Module
