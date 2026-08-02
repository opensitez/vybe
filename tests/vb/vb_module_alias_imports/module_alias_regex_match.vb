' vybe-test: vb/vb_module_alias_imports/module_alias_regex_match
' origin: languages/vb/tests/vb/test_vb_module_alias_imports.rs

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

Imports RegexAlias = System.Text.RegularExpressions.Regex

Module M
    Sub Main()
        __Check(CStr(RegexAlias.IsMatch("abc-123", "^[a-z]+-\d+$")), "True")
        __Check(CStr(RegexAlias.Replace("a-b-c", "-", "_")), "a_b_c")
    End Sub
End Module
