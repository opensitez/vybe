' vybe-test: vb/vb_system_regex_advanced/regex_replace_with_match_evaluator_transforms_matches
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
        Dim result As String = Regex.Replace(
            "a1b2c3",
            "\d",
            Function(m As Match) (CInt(m.Value) * 2).ToString()
        )
        __Check(CStr(result), "a2b4c6")
    End Sub
End Module
