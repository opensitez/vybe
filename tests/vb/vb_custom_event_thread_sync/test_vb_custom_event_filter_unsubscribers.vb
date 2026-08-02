' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_filter_unsubscribers
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

Imports System

Class FilteredBroadcaster
    Private multicast As EventHandler

    Public Custom Event FilteredEvent As EventHandler
        AddHandler(value As EventHandler)
            multicast = CType(Delegate.Combine(multicast, value), EventHandler)
        End AddHandler
        RemoveHandler(value As EventHandler)
            multicast = CType(Delegate.Remove(multicast, value), EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            If multicast IsNot Nothing Then
                Dim invocationList = multicast.GetInvocationList()
                For Each d In invocationList
                    CType(d, EventHandler)(sender, e)
                Next
            End If
        End RaiseEvent
    End Event

    Public Sub Broadcast()
        RaiseEvent FilteredEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim fb As New FilteredBroadcaster()
        AddHandler fb.FilteredEvent, Sub(s, e) Console.WriteLine("Broadcast Received")
        fb.Broadcast()
    End Sub
End Module
