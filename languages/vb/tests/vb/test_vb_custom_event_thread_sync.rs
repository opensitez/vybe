use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Custom Event Thread SyncLock Accessors & Dispatch
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_custom_event_synclock_thread_safe_raise() {
    let src = r#"
Imports System
Imports System.Threading

Class ThreadSafeNotifier
    Private lockObj As New Object()
    Private handlerDelegate As EventHandler

    Public Custom Event StatusUpdate As EventHandler
        AddHandler(value As EventHandler)
            SyncLock lockObj
                handlerDelegate = CType(Delegate.Combine(handlerDelegate, value), EventHandler)
            End SyncLock
        End AddHandler
        RemoveHandler(value As EventHandler)
            SyncLock lockObj
                handlerDelegate = CType(Delegate.Remove(handlerDelegate, value), EventHandler)
            End SyncLock
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            Dim temp As EventHandler
            SyncLock lockObj
                temp = handlerDelegate
            End SyncLock
            If temp IsNot Nothing Then temp(sender, e)
        End RaiseEvent
    End Event

    Public Sub Signal()
        RaiseEvent StatusUpdate(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim notifier As New ThreadSafeNotifier()
        AddHandler notifier.StatusUpdate, Sub(s, e) Console.WriteLine("ThreadSafe Update Received")
        notifier.Signal()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ThreadSafe Update Received"]);
}

#[test]
fn test_vb_custom_event_async_event_raising() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Class AsyncEventSource
    Public Event AsyncNotice As EventHandler

    Public Async Function FireAsync() As Task
        Await Task.Yield()
        RaiseEvent AsyncNotice(Me, EventArgs.Empty)
    End Function
End Class

Module Program
    Sub Main()
        Dim src As New AsyncEventSource()
        AddHandler src.AsyncNotice, Sub(s, e) Console.WriteLine("Async Notice Fired")
        Dim t = src.FireAsync()
        t.Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Async Notice Fired"]);
}

#[test]
fn test_vb_custom_event_handler_multithreaded_subscription() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Concurrent Handlers: 4"]);
}

#[test]
fn test_vb_custom_event_interlocked_exchange_accessor() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Interlocked Event Fired"]);
}

#[test]
fn test_vb_custom_event_task_completion_source_wait() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Class TaskCompletionPublisher
    Public Event TaskCompleted As EventHandler

    Public Sub RunWork()
        RaiseEvent TaskCompleted(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Boolean)()
        Dim pub As New TaskCompletionPublisher()

        AddHandler pub.TaskCompleted, Sub(s, e) tcs.SetResult(True)
        pub.RunWork()

        Console.WriteLine("Task Result: " & tcs.Task.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Task Result: True"]);
}

#[test]
fn test_vb_event_handler_background_thread_execution() {
    let src = r#"
Imports System
Imports System.Threading

Class BackgroundWorkerNotifier
    Public Event WorkDone As EventHandler

    Public Sub StartBackgroundJob()
        Dim t As New Thread(Sub()
            Thread.Sleep(10)
            RaiseEvent WorkDone(Me, EventArgs.Empty)
        End Sub)
        t.Start()
        t.Join()
    End Sub
End Class

Module Program
    Sub Main()
        Dim bwn As New BackgroundWorkerNotifier()
        AddHandler bwn.WorkDone, Sub(s, e) Console.WriteLine("Done on Thread: " & Thread.CurrentThread.IsBackground)
        bwn.StartBackgroundJob()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Done on Thread: True"]);
}

#[test]
fn test_vb_event_handler_synchronization_context_marshal() {
    let src = r#"
Imports System
Imports System.Threading

Class ContextPublisher
    Public Event ContextNotice As EventHandler
    Public Sub Fire()
        RaiseEvent ContextNotice(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim ctx = SynchronizationContext.Current
        Dim pub As New ContextPublisher()
        AddHandler pub.ContextNotice, Sub(s, e) Console.WriteLine("Handled on Current Context")
        pub.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Handled on Current Context"]);
}

#[test]
fn test_vb_custom_event_filter_unsubscribers() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Broadcast Received"]);
}

#[test]
fn test_vb_custom_event_weak_reference_handlers() {
    let src = r#"
Imports System

Class WeakEventSubscriber
    Public Sub OnNotify(sender As Object, e As EventArgs)
        Console.WriteLine("Weak Handler Triggered")
    End Sub
End Class

Class WeakPublisher
    Public Event Notify As EventHandler
    Public Sub Fire()
        RaiseEvent Notify(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New WeakPublisher()
        Dim subObj As New WeakEventSubscriber()
        AddHandler pub.Notify, AddressOf subObj.OnNotify
        pub.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Weak Handler Triggered"]);
}

#[test]
fn test_vb_event_handler_fire_and_forget_async() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Class FireForgetSource
    Public Event WorkTriggered As EventHandler

    Public Sub Fire()
        Task.Run(Sub() RaiseEvent WorkTriggered(Me, EventArgs.Empty)).Wait()
    End Sub
End Class

Module Program
    Sub Main()
        Dim ff As New FireForgetSource()
        AddHandler ff.WorkTriggered, Sub(s, e) Console.WriteLine("Work Triggered Async")
        ff.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Work Triggered Async"]);
}

#[test]
fn test_vb_event_handler_unhandled_exception_in_handler() {
    let src = r#"
Imports System

Class CrashPublisher
    Public Event CrashEvent As EventHandler
    Public Sub Trigger()
        RaiseEvent CrashEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim cp As New CrashPublisher()
        AddHandler cp.CrashEvent, Sub(s, e) Throw New InvalidOperationException("Handler Exception")
        Try
            cp.Trigger()
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Handler Exception"]);
}

#[test]
fn test_vb_custom_event_read_only_accessors() {
    let src = r#"
Imports System

Class ReadOnlyEventSource
    Private list As EventHandler

    Public Custom Event SimpleEvent As EventHandler
        AddHandler(value As EventHandler)
            list = CType(Delegate.Combine(list, value), EventHandler)
        End AddHandler
        RemoveHandler(value As EventHandler)
            list = CType(Delegate.Remove(list, value), EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            If list IsNot Nothing Then list(sender, e)
        End RaiseEvent
    End Event

    Public Sub Run()
        RaiseEvent SimpleEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim r As New ReadOnlyEventSource()
        AddHandler r.SimpleEvent, Sub(s, e) Console.WriteLine("Simple Event Fired")
        r.Run()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Simple Event Fired"]);
}

#[test]
fn test_vb_custom_event_cancel_event_args() {
    let src = r#"
Imports System
Imports System.ComponentModel

Class CancelablePublisher
    Public Event ProcessExecuting As EventHandler(Of CancelEventArgs)

    Public Function ExecuteProcess() As Boolean
        Dim args As New CancelEventArgs()
        RaiseEvent ProcessExecuting(Me, args)
        Return Not args.Cancel
    End Function
End Class

Module Program
    Sub Main()
        Dim cp As New CancelablePublisher()
        AddHandler cp.ProcessExecuting, Sub(s, e) e.Cancel = True
        Dim success = cp.ExecuteProcess()
        Console.WriteLine("Process Allowed: " & success)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Process Allowed: False"]);
}

#[test]
fn test_vb_event_handler_multiple_parallel_triggers() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Class ParallelEventSource
    Public Event Ping As EventHandler
    Public Sub Fire()
        RaiseEvent Ping(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim pes As New ParallelEventSource()
        Dim counter = 0
        Dim lockObj As New Object()
        AddHandler pes.Ping, Sub(s, e)
            SyncLock lockObj
                counter += 1
            End SyncLock
        End Sub

        Parallel.For(0, 5, Sub(i) pes.Fire())
        Console.WriteLine("Parallel Count: " & counter)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Parallel Count: 5"]);
}

#[test]
fn test_vb_custom_event_raise_with_custom_sender() {
    let src = r#"
Imports System

Class VirtualSender
End Class

Class CustomSenderBroadcaster
    Public Event Notice As EventHandler
    Public Sub FireWithSender(customSender As Object)
        RaiseEvent Notice(customSender, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim csb As New CustomSenderBroadcaster()
        Dim virt As New VirtualSender()
        AddHandler csb.Notice, Sub(s, e) Console.WriteLine("Sender Type: " & s.GetType().Name)
        csb.FireWithSender(virt)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Sender Type: VirtualSender"]);
}

#[test]
fn test_vb_event_handler_chained_raise_event() {
    let src = r#"
Imports System

Class EventChain
    Public Event Stage1 As EventHandler
    Public Event Stage2 As EventHandler

    Public Sub StartChain()
        RaiseEvent Stage1(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim ec As New EventChain()
        AddHandler ec.Stage1, Sub(s, e)
            Console.WriteLine("Stage 1")
            ' Raise stage 2 inside stage 1 handler
            AddHandler ec.Stage2, Sub(s2, e2) Console.WriteLine("Stage 2")
        End Sub
        ec.StartChain()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Stage 1"]);
}

#[test]
fn test_vb_event_handler_disposed_publisher_raises_throws() {
    let src = r#"
Imports System

Class DisposablePublisher
    Implements IDisposable
    Public Event DataEvent As EventHandler
    Private isDisposed As Boolean = False

    Public Sub Fire()
        If isDisposed Then Throw New ObjectDisposedException("DisposablePublisher")
        RaiseEvent DataEvent(Me, EventArgs.Empty)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        isDisposed = True
    End Sub
End Class

Module Program
    Sub Main()
        Dim dp As New DisposablePublisher()
        dp.Dispose()
        Try
            dp.Fire()
        Catch ex As ObjectDisposedException
            Console.WriteLine("ObjectDisposedException Caught on Fire")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ObjectDisposedException Caught on Fire"]);
}

#[test]
fn test_vb_custom_event_reentrant_addhandler_removehandler() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Handler 1"]);
}

#[test]
fn test_vb_event_handler_passing_null_event_args() {
    let src = r#"
Imports System

Class NullArgsPublisher
    Public Event RawEvent As EventHandler
    Public Sub FireNullArgs()
        RaiseEvent RawEvent(Me, Nothing)
    End Sub
End Class

Module Program
    Sub Main()
        Dim nap As New NullArgsPublisher()
        AddHandler nap.RawEvent, Sub(s, e) Console.WriteLine(e Is Nothing)
        nap.FireNullArgs()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_custom_event_static_accessor_methods() {
    let src = r#"
Imports System

Class SharedEventSource
    Private Shared handlers As EventHandler

    Public Shared Custom Event SharedNotice As EventHandler
        AddHandler(value As EventHandler)
            handlers = CType(Delegate.Combine(handlers, value), EventHandler)
        End AddHandler
        RemoveHandler(value As EventHandler)
            handlers = CType(Delegate.Remove(handlers, value), EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            If handlers IsNot Nothing Then handlers(sender, e)
        End RaiseEvent
    End Event

    Public Shared Sub Fire()
        RaiseEvent SharedNotice(Nothing, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        AddHandler SharedEventSource.SharedNotice, Sub(s, e) Console.WriteLine("Shared Custom Event Fired")
        SharedEventSource.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Shared Custom Event Fired"]);
}
