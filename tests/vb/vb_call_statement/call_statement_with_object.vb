' vybe-test: vb/vb_call_statement/call_statement_with_object
' origin: languages/vb/tests/vb/test_vb_call_statement.rs

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

Class Logger
    Public Sub Log(msg As String)
        __Check(CStr("Log: " & msg), "Log: Test Call")
    End Sub
End Class

Module M
    Sub Main()
        Dim l As New Logger()
        Call l.Log("Test Call")
    End Sub
End Module
