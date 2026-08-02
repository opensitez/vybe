' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_event_handler_reentrant_event_raising
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

Imports System

Class ReentrantPublisher
    Public Event Ping As EventHandler
    Public Property Count As Integer = 0

    Public Sub Trigger()
        Count += 1
        If Count <= 2 Then
            RaiseEvent Ping(Me, EventArgs.Empty)
        End If
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New ReentrantPublisher()
        AddHandler p.Ping, Sub(s, e)
            Console.WriteLine("Ping " & p.Count)
            p.Trigger()
        End Sub
        p.Trigger()
    End Sub
End Module
