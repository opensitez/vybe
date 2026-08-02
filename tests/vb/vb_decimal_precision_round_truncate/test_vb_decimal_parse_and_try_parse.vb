' vybe-test: vb/vb_decimal_precision_round_truncate/test_vb_decimal_parse_and_try_parse
' origin: languages/vb/tests/vb/test_vb_decimal_precision_round_truncate.rs

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
        Dim val As Decimal
        Dim ok = Decimal.TryParse("1234.56", NumberStyles.Number, CultureInfo.InvariantCulture, val)
        __Check(CStr(ok & "|" & val), "True|1234.56")
    End Sub
End Module
