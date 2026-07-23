use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Constructor Event Subscriptions & Lifecycle Wireup
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_event_subscribed_in_constructor() {
    let src = r#"
Imports System

Class InternalHandler
    Public Property Handled As Boolean = False
    Public Sub New(publisher As EventPublisher)
        AddHandler publisher.Triggered, AddressOf OnPublisherTriggered
    End Sub

    Private Sub OnPublisherTriggered(sender As Object, e As EventArgs)
        Handled = True
    End Sub
End Class

Class EventPublisher
    Public Event Triggered As EventHandler
    Public Sub Fire()
        RaiseEvent Triggered(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New EventPublisher()
        Dim subObj As New InternalHandler(pub)
        pub.Fire()
        Console.WriteLine(subObj.Handled)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_with_events_constructor_wireup() {
    let src = r#"
Imports System

Class Clock
    Public Event Tick As EventHandler
    Public Sub Start()
        RaiseEvent Tick(Me, EventArgs.Empty)
    End Sub
End Class

Class ClockListener
    Private WithEvents myClock As Clock

    Public Sub New(c As Clock)
        myClock = c ' Assigning WithEvents field automatically wires Handles methods!
    End Sub

    Private Sub OnTick(sender As Object, e As EventArgs) Handles myClock.Tick
        Console.WriteLine("Clock Ticked via WithEvents")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Clock()
        Dim listener As New ClockListener(c)
        c.Start()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Clock Ticked via WithEvents"]);
}

#[test]
fn test_vb_with_events_reassignment_unwires_old_instance() {
    let src = r#"
Imports System

Class Emitter
    Public Property Name As String
    Public Event Action As EventHandler
    Public Sub Fire()
        RaiseEvent Action(Me, EventArgs.Empty)
    End Sub
End Class

Class SwitchableListener
    Private WithEvents currentEmitter As Emitter

    Public Sub SetEmitter(e As Emitter)
        currentEmitter = e ' Unwires previous currentEmitter, wires new e!
    End Sub

    Private Sub OnAction(sender As Object, e As EventArgs) Handles currentEmitter.Action
        Console.WriteLine("Action Handled From: " & currentEmitter.Name)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e1 As New Emitter With {.Name = "First"}
        Dim e2 As New Emitter With {.Name = "Second"}

        Dim listener As New SwitchableListener()
        listener.SetEmitter(e1)
        e1.Fire()

        listener.SetEmitter(e2)
        e1.Fire() ' Should NOT fire listener!
        e2.Fire() ' Should fire listener!
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Action Handled From: First", "Action Handled From: Second"]
    );
}

#[test]
fn test_vb_with_events_null_assignment_unwires_handler() {
    let src = r#"
Imports System

Class Emitter
    Public Event Action As EventHandler
    Public Sub Fire()
        RaiseEvent Action(Me, EventArgs.Empty)
    End Sub
End Class

Class NullableListener
    Private WithEvents myEmitter As Emitter

    Public Sub Bind(e As Emitter)
        myEmitter = e
    End Sub

    Private Sub OnAction(sender As Object, e As EventArgs) Handles myEmitter.Action
        Console.WriteLine("Fired")
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        Dim listener As New NullableListener()
        listener.Bind(e)
        e.Fire()

        listener.Bind(Nothing) ' Unwires myEmitter!
        e.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Fired"]);
}

#[test]
fn test_vb_virtual_method_event_handler_subscription_in_constructor() {
    let src = r#"
Imports System

Class BaseSubscriber
    Public Sub New(pub As Source)
        AddHandler pub.Ping, AddressOf OnPing
    End Sub

    Protected Overridable Sub OnPing(sender As Object, e As EventArgs)
        Console.WriteLine("Base OnPing")
    End Sub
End Class

Class DerivedSubscriber
    Inherits BaseSubscriber

    Public Sub New(pub As Source)
        MyBase.New(pub)
    End Sub

    Protected Overrides Sub OnPing(sender As Object, e As EventArgs)
        Console.WriteLine("Derived OnPing")
    End Sub
End Class

Class Source
    Public Event Ping As EventHandler
    Public Sub Fire()
        RaiseEvent Ping(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New Source()
        Dim subObj As New DerivedSubscriber(s)
        s.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Derived OnPing"]);
}

#[test]
fn test_vb_constructor_self_event_subscription() {
    let src = r#"
Imports System

Class SelfNotifyingWidget
    Public Event InternalStateChanged As EventHandler

    Public Sub New()
        AddHandler InternalStateChanged, Sub(s, e) Console.WriteLine("Self Notification Received")
    End Sub

    Public Sub Mutate()
        RaiseEvent InternalStateChanged(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim w As New SelfNotifyingWidget()
        w.Mutate()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Self Notification Received"]);
}

#[test]
fn test_vb_constructor_subscription_lambda_closure() {
    let src = r#"
Imports System

Class LambdaSubscriber
    Public Property SignalCount As Integer = 0
    Public Sub New(pub As Source)
        AddHandler pub.Ping, Sub(s, e) SignalCount += 1
    End Sub
End Class

Class Source
    Public Event Ping As EventHandler
    Public Sub Fire()
        RaiseEvent Ping(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New Source()
        Dim ls As New LambdaSubscriber(s)
        s.Fire()
        s.Fire()
        Console.WriteLine("Signals Received: " & ls.SignalCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Signals Received: 2"]);
}

#[test]
fn test_vb_multiple_withevents_fields_in_same_class() {
    let src = r#"
Imports System

Class EmitterA
    Public Event EventA As EventHandler
    Public Sub Fire()
        RaiseEvent EventA(Me, EventArgs.Empty)
    End Sub
End Class

Class EmitterB
    Public Event EventB As EventHandler
    Public Sub Fire()
        RaiseEvent EventB(Me, EventArgs.Empty)
    End Sub
End Class

Class CombinedListener
    Public WithEvents SourceA As EmitterA
    Public WithEvents SourceB As EmitterB

    Private Sub OnA(sender As Object, e As EventArgs) Handles SourceA.EventA
        Console.WriteLine("A Handled")
    End Sub

    Private Sub OnB(sender As Object, e As EventArgs) Handles SourceB.EventB
        Console.WriteLine("B Handled")
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As New CombinedListener()
        l.SourceA = New EmitterA()
        l.SourceB = New EmitterB()
        l.SourceA.Fire()
        l.SourceB.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A Handled", "B Handled"]);
}

#[test]
fn test_vb_withevents_single_method_handles_multiple_events() {
    let src = r#"
Imports System

Class Source
    Public Event Event1 As EventHandler
    Public Event Event2 As EventHandler
    Public Sub Fire1()
        RaiseEvent Event1(Me, EventArgs.Empty)
    End Sub
    Public Sub Fire2()
        RaiseEvent Event2(Me, EventArgs.Empty)
    End Sub
End Class

Class MultiHandleListener
    Public WithEvents Src As Source

    ' Single handler method handles both Event1 and Event2!
    Private Sub OnCombined(sender As Object, e As EventArgs) Handles Src.Event1, Src.Event2
        Console.WriteLine("Combined Event Handled")
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As New MultiHandleListener With {.Src = New Source()}
        l.Src.Fire1()
        l.Src.Fire2()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Combined Event Handled", "Combined Event Handled"]
    );
}

#[test]
fn test_vb_withevents_property_setter_access_level() {
    let src = r#"
Imports System

Class Publisher
    Public Event Notice As EventHandler
    Public Sub Fire()
        RaiseEvent Notice(Me, EventArgs.Empty)
    End Sub
End Class

Class EncapsulatedListener
    Public WithEvents Pub As Publisher

    Private Sub OnNotice(sender As Object, e As EventArgs) Handles Pub.Notice
        Console.WriteLine("Notice Handled")
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As New EncapsulatedListener()
        Dim p As New Publisher()
        l.Pub = p
        p.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Notice Handled"]);
}

#[test]
fn test_vb_constructor_subscription_weak_reference_simulation() {
    let src = r#"
Imports System

Class ShortLivedSubscriber
    Public Sub New(pub As Broadcaster)
        AddHandler pub.Broadcast, AddressOf HandleBroadcast
    End Sub

    Private Sub HandleBroadcast(sender As Object, e As EventArgs)
        Console.WriteLine("ShortLived Handled")
    End Sub
End Class

Class Broadcaster
    Public Event Broadcast As EventHandler
    Public Sub Fire()
        RaiseEvent Broadcast(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim b As New Broadcaster()
        Dim subObj As New ShortLivedSubscriber(b)
        b.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ShortLived Handled"]);
}

#[test]
fn test_vb_constructor_subscription_unsubscribing_in_dispose() {
    let src = r#"
Imports System

Class ManagedSubscriber
    Implements IDisposable
    Private publisher As Publisher

    Public Sub New(p As Publisher)
        publisher = p
        AddHandler publisher.Notice, AddressOf OnNotice
    End Sub

    Private Sub OnNotice(sender As Object, e As EventArgs)
        Console.WriteLine("Managed Notice")
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If publisher IsNot Nothing Then
            RemoveHandler publisher.Notice, AddressOf OnNotice
            publisher = Nothing
        End If
    End Sub
End Class

Class Publisher
    Public Event Notice As EventHandler
    Public Sub Fire()
        RaiseEvent Notice(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Publisher()
        Using subObj As New ManagedSubscriber(p)
            p.Fire()
        End Using
        p.Fire() ' Should NOT output after dispose!
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Managed Notice"]);
}

#[test]
fn test_vb_withevents_in_base_class_handles_derived_events() {
    let src = r#"
Imports System

Class EventSource
    Public Event Alert As EventHandler
    Public Sub Trigger()
        RaiseEvent Alert(Me, EventArgs.Empty)
    End Sub
End Class

Class BaseListener
    Protected WithEvents Source As EventSource

    Public Sub New(s As EventSource)
        Source = s
    End Sub

    Private Sub OnAlert(sender As Object, e As EventArgs) Handles Source.Alert
        Console.WriteLine("Base Listener Handled Alert")
    End Sub
End Class

Class DerivedListener
    Inherits BaseListener

    Public Sub New(s As EventSource)
        MyBase.New(s)
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New EventSource()
        Dim dl As New DerivedListener(s)
        s.Trigger()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Base Listener Handled Alert"]);
}

#[test]
fn test_vb_withevents_array_not_supported_uses_addhandler() {
    let src = r#"
Imports System

Class Button
    Public Property ID As Integer
    Public Event Click As EventHandler
    Public Sub PerformClick()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Class FormContainer
    Private buttons As Button()

    Public Sub New()
        buttons = New Button() {New Button With {.ID = 1}, New Button With {.ID = 2}}
        For Each btn In buttons
            AddHandler btn.Click, AddressOf OnButtonClick
        Next
    End Sub

    Private Sub OnButtonClick(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        Console.WriteLine("Button " & btn.ID & " Clicked")
    End Sub

    Public Sub TestClicks()
        buttons(0).PerformClick()
        buttons(1).PerformClick()
    End Sub
End Class

Module Program
    Sub Main()
        Dim form As New FormContainer()
        form.TestClicks()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Button 1 Clicked", "Button 2 Clicked"]);
}

#[test]
fn test_vb_constructor_subscription_exception_during_wireup() {
    let src = r#"
Imports System

Class FaultyPublisher
    Public Custom Event CustomEvent As EventHandler
        AddHandler(value As EventHandler)
            Throw New InvalidOperationException("AddHandler Exception")
        End AddHandler
        RemoveHandler(value As EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
        End RaiseEvent
    End Event
End Class

Class FaultySubscriber
    Public Sub New(pub As FaultyPublisher)
        Try
            AddHandler pub.CustomEvent, Sub(s, e)
        Catch ex As InvalidOperationException
            Console.WriteLine("Caught Exception During Wireup")
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Dim fp As New FaultyPublisher()
        Dim fs As New FaultySubscriber(fp)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Caught Exception During Wireup"]);
}

#[test]
fn test_vb_withevents_property_custom_getter_setter() {
    let src = r#"
Imports System

Class Notifier
    Public Event Ping As EventHandler
    Public Sub Fire()
        RaiseEvent Ping(Me, EventArgs.Empty)
    End Sub
End Class

Class ExplicitPropertyListener
    Private _notifier As Notifier

    Public Custom WithEvents Property NotifierProp As Notifier
        Get
            Return _notifier
        End Get
        Set(value As Notifier)
            If _notifier IsNot Nothing Then
                RemoveHandler _notifier.Ping, AddressOf OnPing
            End If
            _notifier = value
            If _notifier IsNot Nothing Then
                AddHandler _notifier.Ping, AddressOf OnPing
            End If
        End Set
    End Property

    Private Sub OnPing(sender As Object, e As EventArgs)
        Console.WriteLine("Explicit Property Handled Ping")
    End Sub
End Class

Module Program
    Sub Main()
        Dim n As New Notifier()
        Dim listener As New ExplicitPropertyListener()
        listener.NotifierProp = n
        n.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Explicit Property Handled Ping"]);
}

#[test]
fn test_vb_constructor_subscription_chaining_events() {
    let src = r#"
Imports System

Class ComponentA
    Public Event EventA As EventHandler
    Public Sub TriggerA()
        RaiseEvent EventA(Me, EventArgs.Empty)
    End Sub
End Class

Class ComponentB
    Public Event EventB As EventHandler
    Public Sub New(compA As ComponentA)
        AddHandler compA.EventA, Sub(s, e) RaiseEvent EventB(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim ca As New ComponentA()
        Dim cb As New ComponentB(ca)
        AddHandler cb.EventB, Sub(s, e) Console.WriteLine("Chained B Handled")
        ca.TriggerA()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Chained B Handled"]);
}

#[test]
fn test_vb_withevents_structure_not_supported_uses_class() {
    let src = r#"
Imports System

Class StructEventSource
    Public Event Signal As EventHandler
    Public Sub Fire()
        RaiseEvent Signal(Me, EventArgs.Empty)
    End Sub
End Class

Class Controller
    Public WithEvents Source As StructEventSource

    Private Sub OnSignal(sender As Object, e As EventArgs) Handles Source.Signal
        Console.WriteLine("Signal Processed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Controller With {.Source = New StructEventSource()}
        c.Source.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Signal Processed"]);
}

#[test]
fn test_vb_constructor_subscription_shared_static_event() {
    let src = r#"
Imports System

Class GlobalNotifier
    Public Shared Event GlobalPing As EventHandler
    Public Shared Sub Fire()
        RaiseEvent GlobalPing(Nothing, EventArgs.Empty)
    End Sub
End Class

Class Subscriber
    Public Sub New()
        AddHandler GlobalNotifier.GlobalPing, AddressOf OnGlobalPing
    End Sub

    Private Sub OnGlobalPing(sender As Object, e As EventArgs)
        Console.WriteLine("Global Ping Received")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New Subscriber()
        GlobalNotifier.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Global Ping Received"]);
}

#[test]
fn test_vb_withevents_subclass_overrides_handler_method() {
    let src = r#"
Imports System

Class Publisher
    Public Event Trigger As EventHandler
    Public Sub Fire()
        RaiseEvent Trigger(Me, EventArgs.Empty)
    End Sub
End Class

Class BaseListener
    Public WithEvents Pub As Publisher

    Protected Overridable Sub OnTrigger(sender As Object, e As EventArgs) Handles Pub.Trigger
        Console.WriteLine("Base OnTrigger")
    End Sub
End Class

Class OverridingListener
    Inherits BaseListener

    Protected Overrides Sub OnTrigger(sender As Object, e As EventArgs)
        Console.WriteLine("Overridden OnTrigger")
    End Sub
End Class

Module Program
    Sub Main()
        Dim ol As New OverridingListener()
        Dim p As New Publisher()
        ol.Pub = p
        p.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Overridden OnTrigger"]);
}
