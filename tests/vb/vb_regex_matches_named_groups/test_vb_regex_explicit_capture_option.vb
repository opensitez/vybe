' vybe-test: vb/vb_regex_matches_named_groups/test_vb_regex_explicit_capture_option
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
        ' ExplicitCapture means un-named (group) is NOT captured!
        Dim regex As New Regex("(unnamed) (?<named>val)", RegexOptions.ExplicitCapture)
        Dim m = regex.Match("unnamed val")
        __Check(CStr(m.Groups.Count & "|" & m.Groups("named").Value), "2|val")
    End Sub
End Module
