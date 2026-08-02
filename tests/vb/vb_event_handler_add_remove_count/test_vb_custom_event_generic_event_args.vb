' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_custom_event_generic_event_args
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

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

Class DataEventArgs
    Inherits EventArgs
    Public Property Payload As String
End Class

Class DataBroadcaster
    Public Event DataReceived As EventHandler(Of DataEventArgs)

    Public Sub Broadcast(data As String)
        RaiseEvent DataReceived(Me, New DataEventArgs With {.Payload = data})
    End Sub
End Class

Module Program
    Sub Main()
        Dim b As New DataBroadcaster()
        AddHandler b.DataReceived, Sub(s, e) __Check(CStr("Data: " & e.Payload), "Data: Payload123")
        b.Broadcast("Payload123")
    End Sub
End Module
