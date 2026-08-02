' vybe-test: vb/vb_string_compare_ordinal_culture/test_vb_string_equality_operator_case_sensitivity
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

Option Compare Binary

Module Program
    Sub Main()
        Dim s1 As String = "abc"
        Dim s2 As String = "ABC"
        __Check(CStr(s1 = s2), "False")
    End Sub
End Module
