' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_currency_exchange_converter
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

Class CurrencyConverter
    Private rates As New Dictionary(Of String, Decimal) From {
        {"USD_EUR", 0.92D},
        {"USD_GBP", 0.78D}
    }

    Public Function Convert(amount As Decimal, pair As String) As Decimal
        Return amount * rates(pair)
    End Function
End Class

Module Program
    Sub Main()
        Dim cc As New CurrencyConverter()
        __Check(CStr(cc.Convert(100D, "USD_EUR")), "92.00")
    End Sub
End Module
