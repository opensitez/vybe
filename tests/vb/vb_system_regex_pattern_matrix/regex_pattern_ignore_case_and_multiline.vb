' vybe-test: vb/vb_system_regex_pattern_matrix/regex_pattern_ignore_case_and_multiline
' origin: languages/vb/tests/vb/test_vb_system_regex_pattern_matrix.rs

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

Module M
    Sub Main()
        Dim input As String = "Line1\nline2\nLINE3"
        Dim matches As MatchCollection = Regex.Matches(input, "line", RegexOptions.IgnoreCase Or RegexOptions.Multiline)

        __Check(CStr(matches.Count), "3")
        __Check(CStr(matches(0).Value), "Line")
    End Sub
End Module
