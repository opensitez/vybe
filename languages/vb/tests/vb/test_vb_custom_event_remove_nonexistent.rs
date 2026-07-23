use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Custom Events RemoveHandler Edge Cases & Handlers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_event_remove_nonexistent_handler_no_op() {
    let src = r#"
Imports System

Class EventSource
    Public Event Action As Action
    Public Sub Fire()
        RaiseEvent Action()
    End Sub
End Class

Module Program
    Private Sub Handler()
        Console.WriteLine("Handler Executed")
    End Sub

    Sub Main()
        Dim src As New EventSource()
        RemoveHandler src.Action, AddressOf Handler
        Console.WriteLine("Remove completed safely")
        src.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Remove completed safely"]);
}

#[test]
fn test_vb_event_remove_lambda_without_reference_no_op() {
    let src = r#"
Imports System

Class Emitter
    Public Event Message As Action(Of String)
    Public Sub Dispatch(m As String)
        RaiseEvent Message(m)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        AddHandler e.Message, Sub(m) Console.WriteLine("Msg1: " & m)
        ' Attempt to remove a different lambda instance with identical body
        RemoveHandler e.Message, Sub(m) Console.WriteLine("Msg1: " & m)
        e.Dispatch("Test")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Msg1: Test"]);
}

#[test]
fn test_vb_event_remove_stored_lambda_reference_succeeds() {
    let src = r#"
Imports System

Class Emitter
    Public Event Message As Action(Of String)
    Public Sub Dispatch(m As String)
        RaiseEvent Message(m)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        Dim handler As Action(Of String) = Sub(m) Console.WriteLine("Msg: " & m)
        AddHandler e.Message, handler
        e.Dispatch("First")
        RemoveHandler e.Message, handler
        e.Dispatch("Second")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Msg: First"]);
}

#[test]
fn test_vb_event_remove_one_of_multiple_subscribers() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Listener 2"]);
}

#[test]
fn test_vb_event_remove_duplicate_subscribers_one_by_one() {
    let src = r#"
Imports System

Class Publisher
    Public Event Tick As Action
    Public Sub Sound()
        RaiseEvent Tick()
    End Sub
End Class

Module Program
    Private Sub OnTick() : Console.WriteLine("Tick") : End Sub

    Sub Main()
        Dim p As New Publisher()
        AddHandler p.Tick, AddressOf OnTick
        AddHandler p.Tick, AddressOf OnTick
        p.Sound()
        Console.WriteLine("---")
        RemoveHandler p.Tick, AddressOf OnTick
        p.Sound()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Tick", "Tick", "---", "Tick"]);
}

#[test]
fn test_vb_event_custom_event_ignore_nonexistent_removal() {
    let src = r#"
Imports System

Class CustomManager
    Private handlers As Action

    Public Custom Event Work As Action
        AddHandler(value As Action)
            handlers = CType([Delegate].Combine(handlers, value), Action)
        End AddHandler
        RemoveHandler(value As Action)
            Dim newHandlers = CType([Delegate].Remove(handlers, value), Action)
            If newHandlers Is Nothing AndAlso handlers IsNot Nothing Then
                Console.WriteLine("All Handlers Removed")
            End If
            handlers = newHandlers
        End RemoveHandler
        RaiseEvent()
            If handlers IsNot Nothing Then handlers()
        End RaiseEvent
    End Event

    Public Sub Run()
        RaiseEvent Work()
    End Sub
End Class

Module Program
    Private Sub Sub1() : Console.WriteLine("Sub1") : End Sub

    Sub Main()
        Dim cm As New CustomManager()
        AddHandler cm.Work, AddressOf Sub1
        RemoveHandler cm.Work, AddressOf Sub1
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["All Handlers Removed"]);
}

#[test]
fn test_vb_event_remove_handler_during_event_execution() {
    let src = r#"
Imports System

Class IterativeEmitter
    Public Event StepEvent As Action
    Public Sub Fire()
        RaiseEvent StepEvent()
    End Sub
End Class

Module Program
    Private h1 As Action
    Private e As New IterativeEmitter()

    Sub Main()
        h1 = Sub()
            Console.WriteLine("H1 Executing & Unsubscribing")
            RemoveHandler e.StepEvent, h1
        End Sub

        AddHandler e.StepEvent, h1
        e.Fire()
        Console.WriteLine("Second Fire:")
        e.Fire()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["H1 Executing & Unsubscribing", "Second Fire:"]
    );
}

#[test]
fn test_vb_event_add_handler_during_event_execution() {
    let src = r#"
Imports System

Class DynamicEmitter
    Public Event Trigger As Action
    Public Sub Fire()
        RaiseEvent Trigger()
    End Sub
End Class

Module Program
    Sub Main()
        Dim de As New DynamicEmitter()
        Dim h2 As Action = Sub() Console.WriteLine("H2 Executed")

        AddHandler de.Trigger, Sub()
            Console.WriteLine("H1 Executing & Adding H2")
            AddHandler de.Trigger, h2
        End Sub

        de.Fire()
        Console.WriteLine("Second Fire:")
        de.Fire()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec![
            "H1 Executing & Adding H2",
            "Second Fire:",
            "H1 Executing & Adding H2",
            "H2 Executed"
        ]
    );
}

#[test]
fn test_vb_event_remove_all_handlers_manually() {
    let src = r#"
Imports System

Class MultiSubscriber
    Public Event Action As Action
    Public Sub Fire()
        RaiseEvent Action()
    End Sub
End Class

Module Program
    Private Sub A() : End Sub
    Private Sub B() : End Sub

    Sub Main()
        Dim ms As New MultiSubscriber()
        AddHandler ms.Action, AddressOf A
        AddHandler ms.Action, AddressOf B
        RemoveHandler ms.Action, AddressOf A
        RemoveHandler ms.Action, AddressOf B
        Console.WriteLine("All removed safely")
        ms.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["All removed safely"]);
}

#[test]
fn test_vb_event_remove_handler_null_target_safe() {
    let src = r#"
Imports System

Class NullTargetEmitter
    Public Event OnData As Action(Of Integer)
    Public Sub Push(v As Integer)
        RaiseEvent OnData(v)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New NullTargetEmitter()
        Dim nullDelegate As Action(Of Integer) = Nothing
        RemoveHandler e.OnData, nullDelegate
        Console.WriteLine("Safely handled null delegate removal")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Safely handled null delegate removal"]);
}

#[test]
fn test_vb_event_remove_handler_interface_reference() {
    let src = r#"
Imports System

Interface IEventContainer
    Event Alert As Action
End Interface

Class ConcreteContainer
    Implements IEventContainer
    Public Event Alert As Action Implements IEventContainer.Alert
    Public Sub Fire()
        RaiseEvent Alert()
    End Sub
End Class

Module Program
    Private Sub OnAlert() : Console.WriteLine("Alerted") : End Sub

    Sub Main()
        Dim c As New ConcreteContainer()
        Dim ic As IEventContainer = c
        AddHandler ic.Alert, AddressOf OnAlert
        c.Fire()
        RemoveHandler ic.Alert, AddressOf OnAlert
        c.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alerted"]);
}

#[test]
fn test_vb_event_custom_event_log_all_operations() {
    let src = r#"
Imports System

Class LoggingPublisher
    Public Custom Event CustomLog As Action(Of String)
        AddHandler(value As Action(Of String))
            Console.WriteLine("Subscribed")
        End AddHandler
        RemoveHandler(value As Action(Of String))
            Console.WriteLine("Unsubscribed")
        End RemoveHandler
        RaiseEvent(msg As String)
        End RaiseEvent
    End Event
End Class

Module Program
    Sub Main()
        Dim p As New LoggingPublisher()
        Dim h As Action(Of String) = Sub(s) End Sub
        AddHandler p.CustomLog, h
        RemoveHandler p.CustomLog, h
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Subscribed", "Unsubscribed"]);
}

#[test]
fn test_vb_event_remove_handler_with_different_instance_same_method() {
    let src = r#"
Imports System

Class Receiver
    Public Sub HandleEvent()
        Console.WriteLine("Event Received")
    End Sub
End Class

Class Emitter
    Public Event EventFired As Action
    Public Sub Fire()
        RaiseEvent EventFired()
    End Sub
End Class

Module Program
    Sub Main()
        Dim r1 As New Receiver()
        Dim r2 As New Receiver()
        Dim e As New Emitter()

        AddHandler e.EventFired, AddressOf r1.HandleEvent
        ' Attempt to remove using r2 instance
        RemoveHandler e.EventFired, AddressOf r2.HandleEvent
        e.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Event Received"]);
}

#[test]
fn test_vb_event_remove_handler_with_same_instance_same_method() {
    let src = r#"
Imports System

Class Receiver
    Public Sub HandleEvent()
        Console.WriteLine("Event Received")
    End Sub
End Class

Class Emitter
    Public Event EventFired As Action
    Public Sub Fire()
        RaiseEvent EventFired()
    End Sub
End Class

Module Program
    Sub Main()
        Dim r1 As New Receiver()
        Dim e As New Emitter()

        AddHandler e.EventFired, AddressOf r1.HandleEvent
        RemoveHandler e.EventFired, AddressOf r1.HandleEvent
        e.Fire()
        Console.WriteLine("No events fired after removal")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["No events fired after removal"]);
}

#[test]
fn test_vb_event_remove_handler_from_static_event() {
    let src = r#"
Imports System

Class SharedPublisher
    Public Shared Event SharedEvent As Action
    Public Shared Sub Fire()
        RaiseEvent SharedEvent()
    End Sub
End Class

Module Program
    Private Sub OnShared() : Console.WriteLine("Shared Fired") : End Sub

    Sub Main()
        AddHandler SharedPublisher.SharedEvent, AddressOf OnShared
        SharedPublisher.Fire()
        RemoveHandler SharedPublisher.SharedEvent, AddressOf OnShared
        SharedPublisher.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Shared Fired"]);
}

#[test]
fn test_vb_event_custom_event_state_tracking() {
    let src = r#"
Imports System

Class FilteredPublisher
    Private count As Integer = 0
    Public Custom Event ItemAdded As Action
        AddHandler(value As Action)
            count += 1
        End AddHandler
        RemoveHandler(value As Action)
            count -= 1
        End RemoveHandler
        RaiseEvent()
        End RaiseEvent
    End Event
    Public Function GetCount() As Integer
        Return count
    End Function
End Class

Module Program
    Private Sub Dummy() : End Sub

    Sub Main()
        Dim fp As New FilteredPublisher()
        AddHandler fp.ItemAdded, AddressOf Dummy
        AddHandler fp.ItemAdded, AddressOf Dummy
        Console.WriteLine(fp.GetCount())
        RemoveHandler fp.ItemAdded, AddressOf Dummy
        Console.WriteLine(fp.GetCount())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "1"]);
}

#[test]
fn test_vb_event_remove_handler_value_type_struct_publisher() {
    let src = r#"
Imports System

Structure StructPublisher
    Public Event Signal As Action
    Public Sub Fire()
        RaiseEvent Signal()
    End Sub
End Structure

Module Program
    Private Sub OnSignal() : Console.WriteLine("Signal") : End Sub

    Sub Main()
        Dim sp As New StructPublisher()
        AddHandler sp.Signal, AddressOf OnSignal
        sp.Fire()
        RemoveHandler sp.Signal, AddressOf OnSignal
        sp.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Signal"]);
}

#[test]
fn test_vb_event_remove_handler_generic_delegate() {
    let src = r#"
Imports System

Class GenericEmitter(Of T)
    Public Event ValueProcessed As EventHandler(Of T)
    Public Sub Process(v As T)
        RaiseEvent ValueProcessed(Me, v)
    End Sub
End Class

Module Program
    Private Sub OnProcess(sender As Object, e As Integer)
        Console.WriteLine("Processed: " & e)
    End Sub

    Sub Main()
        Dim ge As New GenericEmitter(Of Integer)()
        AddHandler ge.ValueProcessed, AddressOf OnProcess
        ge.Process(42)
        RemoveHandler ge.ValueProcessed, AddressOf OnProcess
        ge.Process(100)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Processed: 42"]);
}

#[test]
fn test_vb_event_custom_event_raise_block_without_subscribers() {
    let src = r#"
Imports System

Class SafeCustomEvent
    Public Custom Event EventTest As Action
        AddHandler(value As Action) : End AddHandler
        RemoveHandler(value As Action) : End RemoveHandler
        RaiseEvent()
            Console.WriteLine("RaiseBlock executed directly")
        End RaiseEvent
    End Event
    Public Sub Fire()
        RaiseEvent EventTest()
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New SafeCustomEvent()
        s.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["RaiseBlock executed directly"]);
}

#[test]
fn test_vb_event_remove_handler_chain_unsubscribes_all() {
    let src = r#"
Imports System

Class Emitter
    Public Event Data As Action(Of Integer)
    Public Sub Push(v As Integer)
        RaiseEvent Data(v)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        Dim h1 As Action(Of Integer) = Sub(v) Console.WriteLine("H1:" & v)
        Dim h2 As Action(Of Integer) = Sub(v) Console.WriteLine("H2:" & v)

        AddHandler e.Data, h1
        AddHandler e.Data, h2
        e.Push(1)
        RemoveHandler e.Data, h1
        RemoveHandler e.Data, h2
        e.Push(2)
        Console.WriteLine("Done")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["H1:1", "H2:1", "Done"]);
}
