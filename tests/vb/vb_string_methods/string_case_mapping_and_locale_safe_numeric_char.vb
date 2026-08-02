' vybe-test: vb/vb_string_methods/string_case_mapping_and_locale_safe_numeric_char
' origin: languages/vb/tests/vb/test_vb_string_methods.rs

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
        __Check(CStr("Hello".ToUpper()), "HELLO")
        __Check(CStr("Hello".ToLower()), "hello")
    End Sub
End Module
