' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_shipping_calculator_decorator_pattern
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

Interface IShippingCost
    Function CalculateCost() As Decimal
End Interface

Class BaseShipping
    Implements IShippingCost
    Public Function CalculateCost() As Decimal Implements IShippingCost.CalculateCost
        Return 5.0D
    End Function
End Class

Class ExpressShippingDecorator
    Implements IShippingCost
    Private baseCost As IShippingCost
    Public Sub New(inner As IShippingCost)
        baseCost = inner
    End Sub
    Public Function CalculateCost() As Decimal Implements IShippingCost.CalculateCost
        Return baseCost.CalculateCost() + 15.0D
    End Function
End Class

Module Program
    Sub Main()
        Dim cost As IShippingCost = New ExpressShippingDecorator(New BaseShipping())
        __Check(CStr(cost.CalculateCost()), "20.0")
    End Sub
End Module
