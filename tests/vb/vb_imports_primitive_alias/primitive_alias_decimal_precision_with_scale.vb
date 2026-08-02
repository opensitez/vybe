' vybe-test: vb/vb_imports_primitive_alias/primitive_alias_decimal_precision_with_scale
' origin: languages/vb/tests/vb/test_vb_imports_primitive_alias.rs

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

Imports Money = System.Decimal

Module M
    Sub Main()
        Dim amount As Money = CDec("12.50")
        Dim tax As Money = Money.Round(amount * CDec("0.1"), 2)
        __Check(CStr(amount.ToString("F2")), "12.50")
        __Check(CStr(tax.ToString("F2")), "1.25")
    End Sub
End Module
