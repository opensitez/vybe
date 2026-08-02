' vybe-test: vb/vb_strings_mid_instr/string_mid_statement_assignment
' origin: languages/vb/tests/vb/test_vb_strings_mid_instr.rs

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
        Dim text As String = "Hello Visual Basic"
        ' Mid can be used as a statement to replace parts of a string
        Mid(text, 7, 6) = "Modern"
        __Check(CStr(text), "Hello Modern Basic")
    End Sub
End Module
