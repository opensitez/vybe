' vybe-test: vb/vb_system_string_api_matrix/string_api_matrix_comparisons_and_empty_checks
' origin: languages/vb/tests/vb/test_vb_system_string_api_matrix.rs

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
        __Check(CStr(String.IsNullOrEmpty("")), "True")
        __Check(CStr(String.IsNullOrWhiteSpace("  ")), "True")
        __Check(CStr(String.IsNullOrWhiteSpace("x")), "False")
        __Check(CStr(String.Equals("a", "A")), "False")
        __Check(CStr(String.Equals("a", "a")), "True")
    End Sub
End Module
