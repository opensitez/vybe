' vybe-test: vb/vb_double_try_parse_cultures/test_vb_double_try_parse_german_comma_decimal_separator
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
        Dim deCulture As New CultureInfo("de-DE")
        Dim val As Double
        ' German culture uses comma ',' as decimal separator and dot '.' as group separator!
        Dim ok = Double.TryParse("1.234,56", NumberStyles.Number, deCulture, val)
        __Check(CStr(ok & "|" & val), "True|1234.56")
    End Sub
End Module
