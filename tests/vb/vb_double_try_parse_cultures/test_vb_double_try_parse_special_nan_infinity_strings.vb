' vybe-test: vb/vb_double_try_parse_cultures/test_vb_double_try_parse_special_nan_infinity_strings
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
        Dim valNaN, valInf As Double
        Dim okNaN = Double.TryParse("NaN", NumberStyles.Any, CultureInfo.InvariantCulture, valNaN)
        Dim okInf = Double.TryParse("Infinity", NumberStyles.Any, CultureInfo.InvariantCulture, valInf)
        __Check(CStr(okNaN & "|" & Double.IsNaN(valNaN) & "|" & okInf & "|" & Double.IsPositiveInfinity(valInf)), "True|True|True|True")
    End Sub
End Module
