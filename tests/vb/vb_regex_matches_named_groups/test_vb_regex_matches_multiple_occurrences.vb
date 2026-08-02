' vybe-test: vb/vb_regex_matches_named_groups/test_vb_regex_matches_multiple_occurrences
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
        Dim input = "cat, dog, bird"
        Dim matches = Regex.Matches(input, "\b\w+\b")
        __Check(CStr(matches.Count & ":" & matches(0).Value & "," & matches(1).Value & "," & matches(2).Value), "3:cat,dog,bird")
    End Sub
End Module
