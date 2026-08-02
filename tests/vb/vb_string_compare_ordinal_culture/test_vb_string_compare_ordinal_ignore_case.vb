' vybe-test: vb/vb_string_compare_ordinal_culture/test_vb_string_compare_ordinal_ignore_case
' origin: languages/vb/tests/vb/test_vb_string_compare_ordinal_culture.rs

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

Module Program
    Sub Main()
        Dim res As Integer = String.Compare("abc", "ABC", StringComparison.OrdinalIgnoreCase)
        __Check(CStr(res = 0), "True")
    End Sub
End Module
