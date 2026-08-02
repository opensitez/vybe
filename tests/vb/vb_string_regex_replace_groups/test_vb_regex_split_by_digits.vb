' vybe-test: vb/vb_string_regex_replace_groups/test_vb_regex_split_by_digits
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
        Dim parts As String() = Regex.Split("one1two2three", "\d")
        __Check(CStr(parts.Length), "3")
        __Check(CStr(parts(0)), "one")
        __Check(CStr(parts(1)), "two")
        __Check(CStr(parts(2)), "three")
    End Sub
End Module
