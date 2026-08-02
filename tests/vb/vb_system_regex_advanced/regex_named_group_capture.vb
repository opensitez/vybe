' vybe-test: vb/vb_system_regex_advanced/regex_named_group_capture
' origin: languages/vb/tests/vb/test_vb_system_regex_advanced.rs

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
        Dim m As Match = Regex.Match("date=2024-06-15", "(?<year>\d{4})-(?<month>\d{2})")
        __Check(CStr(m.Groups("year").Value), "2024")
        __Check(CStr(m.Groups("month").Value), "06")
    End Sub
End Module
