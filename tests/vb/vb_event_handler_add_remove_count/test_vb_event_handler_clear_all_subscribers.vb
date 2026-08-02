' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_event_handler_clear_all_subscribers
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

Imports System

Class ClearablePublisher
    Public Event TaskEvent As EventHandler

    Public Sub ClearSubscribers()
        ' In VB.NET inside class, TaskEventEvent represents delegate!
        TaskEventEvent = Nothing
    End Sub

    Public Sub Trigger()
        RaiseEvent TaskEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New ClearablePublisher()
        AddHandler p.TaskEvent, Sub(s, e) Console.WriteLine("Handler 1")
        p.ClearSubscribers()
        p.Trigger()
        Console.WriteLine("Cleared")
    End Sub
End Module
