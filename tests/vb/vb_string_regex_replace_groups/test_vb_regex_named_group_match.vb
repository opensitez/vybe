' vybe-test: vb/vb_string_regex_replace_groups/test_vb_regex_named_group_match
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
        Dim pattern As String = "(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})"
        Dim m As Match = Regex.Match("2026-07-21", pattern)
        __Check(CStr(m.Groups("year").Value), "2026")
        __Check(CStr(m.Groups("month").Value), "07")
        __Check(CStr(m.Groups("day").Value), "21")
    End Sub
End Module
