' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_order_item_subtotal_calculation
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

Class OrderItem
    Public Property ProductId As String
    Public Property UnitPrice As Decimal
    Public Property Quantity As Integer

    Public ReadOnly Property Subtotal As Decimal
        Get
            Return UnitPrice * Quantity
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim item As New OrderItem With {.ProductId = "P1", .UnitPrice = 19.99D, .Quantity = 3}
        __Check(CStr(item.Subtotal), "59.97")
    End Sub
End Module
