' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_custom_event_logging_add_remove
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

Imports System

Class MonitoredEventSource
    Private internalDelegate As EventHandler

    Public Custom Event MonitoredEvent As EventHandler
        AddHandler(value As EventHandler)
            Console.WriteLine("Subscriber Added")
            internalDelegate = CType(Delegate.Combine(internalDelegate, value), EventHandler)
        End AddHandler
        RemoveHandler(value As EventHandler)
            Console.WriteLine("Subscriber Removed")
            internalDelegate = CType(Delegate.Remove(internalDelegate, value), EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            If internalDelegate IsNot Nothing Then internalDelegate(sender, e)
        End RaiseEvent
    End Event
End Class

Module Program
    Sub Main()
        Dim src As New MonitoredEventSource()
        Dim h As EventHandler = Sub(s, e) Console.WriteLine("Fired")
        AddHandler src.MonitoredEvent, h
        RemoveHandler src.MonitoredEvent, h
    End Sub
End Module
