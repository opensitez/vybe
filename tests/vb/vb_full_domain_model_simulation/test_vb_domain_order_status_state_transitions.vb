' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_order_status_state_transitions
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

Enum OrderStatus
    Created
    Paid
    Shipped
    Cancelled
End Enum

Class Order
    Public Property Status As OrderStatus = OrderStatus.Created

    Public Sub Pay()
        If Status <> OrderStatus.Created Then Throw New InvalidOperationException("Cannot pay order in status " & Status)
        Status = OrderStatus.Paid
    End Sub

    Public Sub Ship()
        If Status <> OrderStatus.Paid Then Throw New InvalidOperationException("Cannot ship unpaid order")
        Status = OrderStatus.Shipped
    End Sub
End Class

Module Program
    Sub Main()
        Dim ord As New Order()
        ord.Pay()
        ord.Ship()
        __Check(CStr(ord.Status.ToString()), "Shipped")
    End Sub
End Module
