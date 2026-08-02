' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_event_handler_subscribing_named_method
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

Class NamedSubscriber
    Public Shared Sub OnEvent(sender As Object, e As EventArgs)
        __Check(CStr("Named Method Handled"), "Named Method Handled")
    End Sub
End Class

Class Emitter
    Public Event Trigger As EventHandler
    Public Sub Fire()
        RaiseEvent Trigger(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim em As New Emitter()
        AddHandler em.Trigger, AddressOf NamedSubscriber.OnEvent
        em.Fire()
    End Sub
End Module
