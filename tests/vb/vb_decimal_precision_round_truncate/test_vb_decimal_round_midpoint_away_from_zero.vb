' vybe-test: vb/vb_decimal_precision_round_truncate/test_vb_decimal_round_midpoint_away_from_zero
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

Module Program
    Sub Main()
        __Check(CStr(Decimal.Round(2.5D, MidpointRounding.AwayFromZero) & "|" & Decimal.Round(3.5D, MidpointRounding.AwayFromZero)), "3|4")
    End Sub
End Module
