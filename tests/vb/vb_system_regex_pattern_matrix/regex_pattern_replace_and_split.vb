' vybe-test: vb/vb_system_regex_pattern_matrix/regex_pattern_replace_and_split
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
        Dim source As String = "a1 b2 c3"
        Dim clean As String = Regex.Replace(source, "\d", "")
        Dim parts As String() = Regex.Split("a,b;;c", "[;,]+")

        __Check(CStr(clean = "a b c"), "True")
        __Check(CStr(parts.Length), "3")
        __Check(CStr(parts(0)), "a")
        __Check(CStr(parts(2)), "c")
    End Sub
End Module
