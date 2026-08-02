' vybe-test: vb/vb_strings_mid_instr/string_instr_function
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
        Dim text As String = "apple banana apple"
        ' InStr returns 1-based index
        __Check(CStr(InStr(text, "banana")), "7")
        ' Start search from index 8
        __Check(CStr(InStr(8, text, "apple")), "14")
        ' Case insensitive search (CompareMethod.Text = 1)
        __Check(CStr(InStr(1, "Hello", "h", 1)), "1")
    End Sub
End Module
