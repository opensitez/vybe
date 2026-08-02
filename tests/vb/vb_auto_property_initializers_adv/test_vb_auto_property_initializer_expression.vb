' vybe-test: vb/vb_auto_property_initializers_adv/test_vb_auto_property_initializer_expression
' origin: languages/vb/tests/vb/test_vb_auto_property_initializers_adv.rs

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

Class ShoppingCart
    Public Property Items As New List(Of String) From {"Item1", "Item2"}
End Class

Module Program
    Sub Main()
        Dim cart As New ShoppingCart()
        __Check(CStr(cart.Items.Count), "2")
        __Check(CStr(cart.Items(0)), "Item1")
    End Sub
End Module
