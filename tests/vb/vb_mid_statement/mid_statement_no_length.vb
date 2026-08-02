' vybe-test: vb/vb_mid_statement/mid_statement_no_length
' origin: languages/vb/tests/vb/test_vb_mid_statement.rs

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
        Dim text As String = "Apple"
        
        ' If length is omitted, it replaces as much as possible
        Mid(text, 2) = "nna"
        __Check(CStr(text), "Annna")
    End Sub
End Module
