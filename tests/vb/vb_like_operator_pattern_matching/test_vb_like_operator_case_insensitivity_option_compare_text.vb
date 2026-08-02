' vybe-test: vb/vb_like_operator_pattern_matching/test_vb_like_operator_case_insensitivity_option_compare_text
' origin: languages/vb/tests/vb/test_vb_like_operator_pattern_matching.rs

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

Option Compare Text

Module Program
    Sub Main()
        __Check(CStr("abc" Like "ABC"), "True")
    End Sub
End Module
