' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_customer_tier_discount_calculation
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

Enum CustomerTier
    Standard
    Silver
    Gold
End Enum

Class Customer
    Public Property Tier As CustomerTier = CustomerTier.Gold

    Public Function CalculateDiscount(total As Decimal) As Decimal
        Select Case Tier
            Case CustomerTier.Gold
                Return total * 0.2D
            Case CustomerTier.Silver
                Return total * 0.1D
            Case Else
                Return 0D
        End Select
    End Function
End Class

Module Program
    Sub Main()
        Dim cust As New Customer()
        __Check(CStr(cust.CalculateDiscount(100D)), "20.0")
    End Sub
End Module
