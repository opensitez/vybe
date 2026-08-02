' vybe-test: vb/vb_regex_matches_named_groups/test_vb_regex_named_group_extraction
' origin: languages/vb/tests/vb/test_vb_regex_matches_named_groups.rs

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
        Dim input = "User: Alice, Age: 30"
        Dim pattern = "User: (?<name>\w+), Age: (?<age>\d+)"
        Dim match = Regex.Match(input, pattern)
        __Check(CStr(match.Groups("name").Value & "|" & match.Groups("age").Value), "Alice|30")
    End Sub
End Module
