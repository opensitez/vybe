' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_custom_event_invocation_list_subscriber_count
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

Imports System

Class Publisher
    Private delegateList As EventHandler

    Public Custom Event StatusChanged As EventHandler
        AddHandler(value As EventHandler)
            delegateList = CType(Delegate.Combine(delegateList, value), EventHandler)
        End AddHandler
        RemoveHandler(value As EventHandler)
            delegateList = CType(Delegate.Remove(delegateList, value), EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            If delegateList IsNot Nothing Then delegateList(sender, e)
        End RaiseEvent
    End Event

    Public Function GetSubscriberCount() As Integer
        Return If(delegateList IsNot Nothing, delegateList.GetInvocationList().Length, 0)
    End Function
End Class

Module Program
    Sub Main()
        Dim p As New Publisher()
        Dim h1 As EventHandler = Sub(s, e) Console.WriteLine("H1")
        Dim h2 As EventHandler = Sub(s, e) Console.WriteLine("H2")

        AddHandler p.StatusChanged, h1
        AddHandler p.StatusChanged, h2
        Console.WriteLine("Subscribers: " & p.GetSubscriberCount())
    End Sub
End Module
