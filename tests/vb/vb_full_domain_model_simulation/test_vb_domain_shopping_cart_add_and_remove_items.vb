' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_shopping_cart_add_and_remove_items
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
Imports System.Linq

Class CartItem
    Public Property Id As String
    Public Property Price As Decimal
End Class

Class ShoppingCart
    Private items As New List(Of CartItem)()

    Public Sub Add(item As CartItem)
        items.Add(item)
    End Sub

    Public Sub Remove(id As String)
        items.RemoveAll(Function(i) i.Id = id)
    End Sub

    Public ReadOnly Property Total As Decimal
        Get
            Return items.Sum(Function(i) i.Price)
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim cart As New ShoppingCart()
        cart.Add(New CartItem With {.Id = "I1", .Price = 10D})
        cart.Add(New CartItem With {.Id = "I2", .Price = 20D})
        cart.Remove("I1")
        __Check(CStr(cart.Total), "20")
    End Sub
End Module
