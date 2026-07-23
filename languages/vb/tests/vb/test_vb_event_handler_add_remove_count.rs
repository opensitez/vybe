use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Custom AddHandler / RemoveHandler / RaiseEvent Accessors
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_custom_event_addhandler_removehandler_accessor() {
    let src = r#"
Imports System

Class Button
    Private handlers As EventHandler

    Public Custom Event Click As EventHandler
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

    Public Sub PerformClick()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim btn As New Button()
        Dim clicked = False
        Dim handler As EventHandler = Sub(s, e) clicked = True

        AddHandler btn.Click, handler
        btn.PerformClick()
        Console.WriteLine(clicked)

        clicked = False
        RemoveHandler btn.Click, handler
        btn.PerformClick()
        Console.WriteLine(clicked)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_custom_event_invocation_list_subscriber_count() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Subscribers: 2"]);
}

#[test]
fn test_vb_custom_event_generic_event_args() {
    let src = r#"
Imports System

Class DataEventArgs
    Inherits EventArgs
    Public Property Payload As String
End Class

Class DataBroadcaster
    Public Event DataReceived As EventHandler(Of DataEventArgs)

    Public Sub Broadcast(data As String)
        RaiseEvent DataReceived(Me, New DataEventArgs With {.Payload = data})
    End Sub
End Class

Module Program
    Sub Main()
        Dim b As New DataBroadcaster()
        AddHandler b.DataReceived, Sub(s, e) Console.WriteLine("Data: " & e.Payload)
        b.Broadcast("Payload123")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Data: Payload123"]);
}

#[test]
fn test_vb_custom_event_prevent_duplicate_handler() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Class UniquePublisher
    Private handlerList As New List(Of EventHandler)()

    Public Custom Event UniqueEvent As EventHandler
        AddHandler(value As EventHandler)
            If Not handlerList.Contains(value) Then handlerList.Add(value)
        End AddHandler
        RemoveHandler(value As EventHandler)
            handlerList.Remove(value)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            For Each h In handlerList
                h(sender, e)
            Next
        End RaiseEvent
    End Event

    Public Sub Trigger()
        RaiseEvent UniqueEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New UniquePublisher()
        Dim count = 0
        Dim handler As EventHandler = Sub(s, e) count += 1

        AddHandler p.UniqueEvent, handler
        AddHandler p.UniqueEvent, handler ' Duplicate add ignored by custom logic!
        p.Trigger()
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_event_handler_multiple_subscribers_order() {
    let src = r#"
Imports System

Class OrderPublisher
    Public Event ActionExecuted As EventHandler

    Public Sub Run()
        RaiseEvent ActionExecuted(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New OrderPublisher()
        AddHandler pub.ActionExecuted, Sub(s, e) Console.WriteLine("Step 1")
        AddHandler pub.ActionExecuted, Sub(s, e) Console.WriteLine("Step 2")
        pub.Run()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Step 1", "Step 2"]);
}

#[test]
fn test_vb_event_handler_null_raise_safe() {
    let src = r#"
Imports System

Class QuietPublisher
    Public Event QuietEvent As EventHandler

    Public Sub Trigger()
        ' RaiseEvent with zero subscribers in standard Event is safe!
        RaiseEvent QuietEvent(Me, EventArgs.Empty)
        Console.WriteLine("Triggered Safely")
    End Sub
End Class

Module Program
    Sub Main()
        Dim q As New QuietPublisher()
        q.Trigger()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Triggered Safely"]);
}

#[test]
fn test_vb_event_handler_subscribing_named_method() {
    let src = r#"
Imports System

Class NamedSubscriber
    Public Shared Sub OnEvent(sender As Object, e As EventArgs)
        Console.WriteLine("Named Method Handled")
    End Sub
End Class

Class Emitter
    Public Event Trigger As EventHandler
    Public Sub Fire()
        RaiseEvent Trigger(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim em As New Emitter()
        AddHandler em.Trigger, AddressOf NamedSubscriber.OnEvent
        em.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Named Method Handled"]);
}

#[test]
fn test_vb_event_handler_removing_named_method() {
    let src = r#"
Imports System

Class Emitter
    Public Event Trigger As EventHandler
    Public Sub Fire()
        RaiseEvent Trigger(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Private Sub OnTrigger(sender As Object, e As EventArgs)
        Console.WriteLine("Triggered")
    End Sub

    Sub Main()
        Dim em As New Emitter()
        AddHandler em.Trigger, AddressOf OnTrigger
        em.Fire()

        RemoveHandler em.Trigger, AddressOf OnTrigger
        em.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Triggered"]);
}

#[test]
fn test_vb_custom_event_logging_add_remove() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Subscriber Added", "Subscriber Removed"]);
}

#[test]
fn test_vb_event_handler_lambda_closure_state_mutation() {
    let src = r#"
Imports System

Class CounterEmitter
    Public Event Increment As EventHandler
    Public Sub Fire()
        RaiseEvent Increment(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim count = 0
        Dim emitter As New CounterEmitter()
        AddHandler emitter.Increment, Sub(s, e) count += 10
        emitter.Fire()
        emitter.Fire()
        Console.WriteLine("Final Count: " & count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Final Count: 20"]);
}

#[test]
fn test_vb_event_handler_custom_delegate_signature() {
    let src = r#"
Delegate Sub CustomStatusHandler(code As Integer, message As String)

Class StatusNotifier
    Public Event StatusReport As CustomStatusHandler
    Public Sub Notify(c As Integer, m As String)
        RaiseEvent StatusReport(c, m)
    End Sub
End Class

Module Program
    Sub Main()
        Dim n As New StatusNotifier()
        AddHandler n.StatusReport, Sub(c, m) Console.WriteLine(c & ": " & m)
        n.Notify(200, "OK")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["200: OK"]);
}

#[test]
fn test_vb_event_handler_raise_in_derived_class() {
    let src = r#"
Imports System

Class BaseNotifier
    Public Event OnNotice As EventHandler
    Protected Sub RaiseNotice()
        RaiseEvent OnNotice(Me, EventArgs.Empty)
    End Sub
End Class

Class DerivedNotifier
    Inherits BaseNotifier
    Public Sub TriggerFromDerived()
        RaiseNotice()
    End Sub
End Class

Module Program
    Sub Main()
        Dim dn As New DerivedNotifier()
        AddHandler dn.OnNotice, Sub(s, e) Console.WriteLine("Notice From Derived")
        dn.TriggerFromDerived()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Notice From Derived"]);
}

#[test]
fn test_vb_custom_event_thread_safe_accessors() {
    let src = r#"
Imports System
Imports System.Threading

Class ThreadSafeEventSource
    Private lockObj As New Object()
    Private handlers As EventHandler

    Public Custom Event SafeEvent As EventHandler
        AddHandler(value As EventHandler)
            SyncLock lockObj
                handlers = CType(Delegate.Combine(handlers, value), EventHandler)
            End SyncLock
        End AddHandler
        RemoveHandler(value As EventHandler)
            SyncLock lockObj
                handlers = CType(Delegate.Remove(handlers, value), EventHandler)
            End SyncLock
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            Dim copy As EventHandler
            SyncLock lockObj
                copy = handlers
            End SyncLock
            If copy IsNot Nothing Then copy(sender, e)
        End RaiseEvent
    End Event

    Public Sub Fire()
        RaiseEvent SafeEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim src As New ThreadSafeEventSource()
        AddHandler src.SafeEvent, Sub(s, e) Console.WriteLine("Thread Safe Event Fired")
        src.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Thread Safe Event Fired"]);
}

#[test]
fn test_vb_event_handler_clear_all_subscribers() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Cleared"]);
}

#[test]
fn test_vb_event_handler_exception_handling_during_raise() {
    let src = r#"
Imports System

Class RobustPublisher
    Public Event ActionEvent As EventHandler

    Public Sub SafeRaise()
        If ActionEventEvent IsNot Nothing Then
            For Each del In ActionEventEvent.GetInvocationList()
                Try
                    del.DynamicInvoke(Me, EventArgs.Empty)
                Catch ex As Exception
                    Console.WriteLine("Handler Error Handled")
                End Try
            Next
        End If
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New RobustPublisher()
        AddHandler p.ActionEvent, Sub(s, e) Throw New InvalidOperationException("Handler Fail")
        AddHandler p.ActionEvent, Sub(s, e) Console.WriteLine("Handler 2 Executed")
        p.SafeRaise()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Handler Error Handled", "Handler 2 Executed"]
    );
}

#[test]
fn test_vb_event_handler_interface_contract() {
    let src = r#"
Imports System

Interface IClickable
    Event Click As EventHandler
End Interface

Class ButtonWidget
    Implements IClickable
    Public Event Click As EventHandler Implements IClickable.Click

    Public Sub ClickMe()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim widget As IClickable = New ButtonWidget()
        AddHandler widget.Click, Sub(s, e) Console.WriteLine("Interface Click Handled")
        CType(widget, ButtonWidget).ClickMe()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Interface Click Handled"]);
}

#[test]
fn test_vb_event_handler_static_shared_event() {
    let src = r#"
Imports System

Class GlobalEvents
    Public Shared Event OnAppExit As EventHandler
    Public Shared Sub TriggerExit()
        RaiseEvent OnAppExit(Nothing, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        AddHandler GlobalEvents.OnAppExit, Sub(s, e) Console.WriteLine("App Exiting")
        GlobalEvents.TriggerExit()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["App Exiting"]);
}

#[test]
fn test_vb_event_handler_value_type_sender() {
    let src = r#"
Imports System

Class ValueSenderNotifier
    Public Event CustomEvent As EventHandler

    Public Sub TriggerValueSender()
        ' Boxed integer sender
        RaiseEvent CustomEvent(100, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim n As New ValueSenderNotifier()
        AddHandler n.CustomEvent, Sub(s, e) Console.WriteLine("Sender Value: " & s.ToString())
        n.TriggerValueSender()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Sender Value: 100"]);
}

#[test]
fn test_vb_event_handler_subscribing_multiple_events_same_handler() {
    let src = r#"
Imports System

Class DualSource
    Public Event EventA As EventHandler
    Public Event EventB As EventHandler

    Public Sub TriggerBoth()
        RaiseEvent EventA(Me, EventArgs.Empty)
        RaiseEvent EventB(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim src As New DualSource()
        Dim commonHandler As EventHandler = Sub(s, e) Console.WriteLine("Common Handler Fired")

        AddHandler src.EventA, commonHandler
        AddHandler src.EventB, commonHandler
        src.TriggerBoth()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Common Handler Fired", "Common Handler Fired"]
    );
}

#[test]
fn test_vb_event_handler_reentrant_event_raising() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Ping 1", "Ping 2"]);
}
