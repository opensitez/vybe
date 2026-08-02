' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_reentrant_addhandler_removehandler
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

Imports System

Class DynamicEventPublisher
    Public Event DynamicEvent As EventHandler

    Public Sub Trigger()
        RaiseEvent DynamicEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim dep As New DynamicEventPublisher()
        Dim h2 As EventHandler = Sub(s, e) Console.WriteLine("Handler 2")
        Dim h1 As EventHandler = Sub(s, e)
            Console.WriteLine("Handler 1")
            AddHandler dep.DynamicEvent, h2
        End Sub

        AddHandler dep.DynamicEvent, h1
        dep.Trigger()
    End Sub
End Module
