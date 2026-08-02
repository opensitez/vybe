' vybe-test: vb/vb_system_regex_static/system_regex_static_is_match_and_match_value
' origin: languages/vb/tests/vb/test_vb_system_regex_static.rs

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
        __Check(CStr(Regex.IsMatch("invoice-42", "\d+")), "True")
        __Check(CStr(Regex.Match("item99", "\d+").Value), "99")
    End Sub
End Module
