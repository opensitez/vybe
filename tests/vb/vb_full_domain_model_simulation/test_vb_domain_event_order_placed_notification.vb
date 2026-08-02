' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_event_order_placed_notification
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

Class OrderPlacedEventArgs
    Inherits EventArgs
    Public Property OrderId As String
End Class

Class OrderService
    Public Event OrderPlaced As EventHandler(Of OrderPlacedEventArgs)

    Public Sub PlaceOrder(id As String)
        RaiseEvent OrderPlaced(Me, New OrderPlacedEventArgs With {.OrderId = id})
    End Sub
End Class

Module Program
    Sub Main()
        Dim service As New OrderService()
        AddHandler service.OrderPlaced, Sub(s, e) __Check(CStr("Notification Sent For Order: " & e.OrderId), "Notification Sent For Order: ORD-999")
        service.PlaceOrder("ORD-999")
    End Sub
End Module
