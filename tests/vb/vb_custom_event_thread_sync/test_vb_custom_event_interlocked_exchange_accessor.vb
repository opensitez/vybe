' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_interlocked_exchange_accessor
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

Imports System
Imports System.Threading

Class InterlockedEventSource
    Private handlers As EventHandler

    Public Custom Event FastEvent As EventHandler
        AddHandler(value As EventHandler)
            Dim oldHandlers As EventHandler = Nothing
            Dim newHandlers As EventHandler = Nothing
            Do
                oldHandlers = handlers
                newHandlers = CType(Delegate.Combine(oldHandlers, value), EventHandler)
            Loop While Interlocked.CompareExchange(handlers, newHandlers, oldHandlers) IsNot oldHandlers
        End AddHandler

        RemoveHandler(value As EventHandler)
            Dim oldHandlers As EventHandler = Nothing
            Dim newHandlers As EventHandler = Nothing
            Do
                oldHandlers = handlers
                newHandlers = CType(Delegate.Remove(oldHandlers, value), EventHandler)
            Loop While Interlocked.CompareExchange(handlers, newHandlers, oldHandlers) IsNot oldHandlers
        End RemoveHandler

        RaiseEvent(sender As Object, e As EventArgs)
            Dim currentHandlers As EventHandler = Volatile.Read(handlers)
            If currentHandlers IsNot Nothing Then currentHandlers(sender, e)
        End RaiseEvent
    End Event

    Public Sub Fire()
        RaiseEvent FastEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim ies As New InterlockedEventSource()
        AddHandler ies.FastEvent, Sub(s, e) Console.WriteLine("Interlocked Event Fired")
        ies.Fire()
    End Sub
End Module
