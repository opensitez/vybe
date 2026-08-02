' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_handler_multithreaded_subscription
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

Imports System
Imports System.Threading
Imports System.Threading.Tasks

Class ConcurrentNotifier
    Private lockObj As New Object()
    Private multicast As EventHandler

    Public Custom Event SharedEvent As EventHandler
        AddHandler(value As EventHandler)
            SyncLock lockObj
                multicast = CType(Delegate.Combine(multicast, value), EventHandler)
            End SyncLock
        End AddHandler
        RemoveHandler(value As EventHandler)
            SyncLock lockObj
                multicast = CType(Delegate.Remove(multicast, value), EventHandler)
            End SyncLock
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            Dim copy As EventHandler
            SyncLock lockObj
                copy = multicast
            End SyncLock
            If copy IsNot Nothing Then copy(sender, e)
        End RaiseEvent
    End Event

    Public Function GetCount() As Integer
        SyncLock lockObj
            Return If(multicast IsNot Nothing, multicast.GetInvocationList().Length, 0)
        End SyncLock
    End Function
End Class

Module Program
    Sub Main()
        Dim cn As New ConcurrentNotifier()
        Dim tasks(3) As Task
        For i As Integer = 0 To 3
            tasks(i) = Task.Run(Sub()
                AddHandler cn.SharedEvent, Sub(s, e)
                End Sub
            End Sub)
        Next
        Task.WaitAll(tasks)
        Console.WriteLine("Concurrent Handlers: " & cn.GetCount())
    End Sub
End Module
