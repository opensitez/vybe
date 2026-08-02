' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_order_tracking_history
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

Class TrackingEvent
    Public Status As String
    Public Location As String
    Public Sub New(s As String, l As String)
        Status = s
        Location = l
    End Sub
End Class

Module Program
    Sub Main()
        Dim history As New Stack(Of TrackingEvent)()
        history.Push(New TrackingEvent("Dispatched", "Warehouse A"))
        history.Push(New TrackingEvent("In Transit", "Hub B"))
        history.Push(New TrackingEvent("Out For Delivery", "Local Hub"))

        Dim current = history.Peek()
        __Check(CStr(current.Status & " at " & current.Location), "Out For Delivery at Local Hub")
    End Sub
End Module
