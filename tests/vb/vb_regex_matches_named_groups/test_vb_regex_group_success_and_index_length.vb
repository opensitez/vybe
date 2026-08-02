' vybe-test: vb/vb_regex_matches_named_groups/test_vb_regex_group_success_and_index_length
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
        Dim match = Regex.Match("ABC 123 XYZ", "\d+")
        __Check(CStr(match.Success & "|" & match.Index & "|" & match.Length), "True|4|3")
    End Sub
End Module
