' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_handler_with_same_instance_same_method
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

Imports System

Class Receiver
    Public Sub HandleEvent()
        Console.WriteLine("Event Received")
    End Sub
End Class

Class Emitter
    Public Event EventFired As Action
    Public Sub Fire()
        RaiseEvent EventFired()
    End Sub
End Class

Module Program
    Sub Main()
        Dim r1 As New Receiver()
        Dim e As New Emitter()

        AddHandler e.EventFired, AddressOf r1.HandleEvent
        RemoveHandler e.EventFired, AddressOf r1.HandleEvent
        e.Fire()
        Console.WriteLine("No events fired after removal")
    End Sub
End Module
