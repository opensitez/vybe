' vybe-test: vb/vb_system_regex_pattern_matrix/regex_pattern_simple_match_capture
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
        Dim pattern As String = "^(?<first>\w+)-(?<second>\d+)$"
        Dim value As String = "item-123"
        Dim m As Match = Regex.Match(value, pattern)

        __Check(CStr(m.Success), "True")
        __Check(CStr(m.Groups("first").Value), "item")
        __Check(CStr(m.Groups("second").Value), "123")
    End Sub
End Module
