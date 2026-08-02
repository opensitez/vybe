' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_tax_calculator_by_region
' origin: languages/vb/tests/vb/test_vb_full_domain_model_simulation.rs

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

Imports System.Collections.Generic

Class TaxCalculator
    Private taxRates As New Dictionary(Of String, Decimal) From {
        {"NY", 0.08D},
        {"CA", 0.10D},
        {"TX", 0.06D}
    }

    Public Function CalculateTax(state As String, amount As Decimal) As Decimal
        If taxRates.ContainsKey(state) Then
            Return amount * taxRates(state)
        End If
        Return 0D
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As New TaxCalculator()
        __Check(CStr(calc.CalculateTax("CA", 200D)), "20.00")
    End Sub
End Module
