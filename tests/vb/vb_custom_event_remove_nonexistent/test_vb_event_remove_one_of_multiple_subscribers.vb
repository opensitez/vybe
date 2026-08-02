' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_one_of_multiple_subscribers
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

Imports System

Class Broadcaster
    Public Event Signal As Action
    Public Sub Send()
        RaiseEvent Signal()
    End Sub
End Class

Module Program
    Private Sub Listener1() : Console.WriteLine("Listener 1") : End Sub
    Private Sub Listener2() : Console.WriteLine("Listener 2") : End Sub

    Sub Main()
        Dim b As New Broadcaster()
        AddHandler b.Signal, AddressOf Listener1
        AddHandler b.Signal, AddressOf Listener2
        RemoveHandler b.Signal, AddressOf Listener1
        b.Send()
    End Sub
End Module
