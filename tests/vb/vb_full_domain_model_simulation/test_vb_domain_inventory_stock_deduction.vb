' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_inventory_stock_deduction
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

Imports System

Class InventoryItem
    Public Property Sku As String
    Public Property QuantityOnHand As Integer

    Public Sub DeductStock(qty As Integer)
        If qty > QuantityOnHand Then Throw New InvalidOperationException("Insufficient stock for " & Sku)
        QuantityOnHand -= qty
    End Sub
End Class

Module Program
    Sub Main()
        Dim inv As New InventoryItem With {.Sku = "SKU-A", .QuantityOnHand = 50}
        inv.DeductStock(10)
        __Check(CStr(inv.QuantityOnHand), "40")
    End Sub
End Module
