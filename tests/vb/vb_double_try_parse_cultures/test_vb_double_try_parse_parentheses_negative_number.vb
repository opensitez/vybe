' vybe-test: vb/vb_double_try_parse_cultures/test_vb_double_try_parse_parentheses_negative_number
' origin: languages/vb/tests/vb/test_vb_double_try_parse_cultures.rs

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

Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim val As Double
        ' AllowParentheses parses (100.5) as -100.5!
        Dim ok = Double.TryParse("(100.5)", NumberStyles.AllowParentheses Or NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, val)
        __Check(CStr(ok & "|" & val), "True|-100.5")
    End Sub
End Module
