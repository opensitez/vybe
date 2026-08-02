' vybe-test: vb/vb_string_regex_replace_groups/test_vb_regex_is_match_timeout
' origin: languages/vb/tests/vb/test_vb_string_regex_replace_groups.rs

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

Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim isMatch As Boolean = Regex.IsMatch("abc", "^a", RegexOptions.None, TimeSpan.FromSeconds(1))
        __Check(CStr(isMatch), "True")
    End Sub
End Module
