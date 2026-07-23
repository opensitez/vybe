use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Events Raising, Invocation List & Multicast Safety
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_event_raise_single_subscriber() {
    let src = r#"
Imports System

Class Button
    Public Event Click As EventHandler
    Public Sub PerformClick()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim b As New Button()
        AddHandler b.Click, Sub(sender, args) Console.WriteLine("Button Clicked")
        b.PerformClick()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Button Clicked"]);
}

#[test]
fn test_vb_event_raise_multiple_subscribers_invocation_order() {
    let src = r#"
Imports System

Class Emitter
    Public Event Trigger As Action(Of Integer)
    Public Sub Fire(val As Integer)
        RaiseEvent Trigger(val)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        AddHandler e.Trigger, Sub(v) Console.WriteLine("Sub1: " & v)
        AddHandler e.Trigger, Sub(v) Console.WriteLine("Sub2: " & (v * 2))
        e.Fire(10)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Sub1: 10", "Sub2: 20"]);
}

#[test]
fn test_vb_event_raise_no_subscribers_safe_noop() {
    let src = r#"
Imports System

Class SafeEmitter
    Public Event SafeEvent As EventHandler
    Public Sub Fire()
        RaiseEvent SafeEvent(Me, EventArgs.Empty)
        Console.WriteLine("Fire Safe Completed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New SafeEmitter()
        e.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Fire Safe Completed"]);
}

#[test]
fn test_vb_event_remove_handler_unsubscribes() {
    let src = r#"
Imports System

Class Alarm
    Public Event Ring As Action
    Public Sub Sound()
        RaiseEvent Ring()
    End Sub
End Class

Module Program
    Private Sub OnRing()
        Console.WriteLine("Alarm Ringing")
    End Sub

    Sub Main()
        Dim a As New Alarm()
        AddHandler a.Ring, AddressOf OnRing
        a.Sound()
        RemoveHandler a.Ring, AddressOf OnRing
        a.Sound()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alarm Ringing"]);
}

#[test]
fn test_vb_event_custom_event_add_remove_raise_blocks() {
    let src = r#"
Imports System

Class CustomPublisher
    Private handlerList As Action(Of String)

    Public Custom Event Message As Action(Of String)
        AddHandler(value As Action(Of String))
            handlerList = CType([Delegate].Combine(handlerList, value), Action(Of String))
            Console.WriteLine("Custom AddHandler")
        End AddHandler
        RemoveHandler(value As Action(Of String))
            handlerList = CType([Delegate].Remove(handlerList, value), Action(Of String))
            Console.WriteLine("Custom RemoveHandler")
        End RemoveHandler
        RaiseEvent(msg As String)
            If handlerList IsNot Nothing Then handlerList(msg)
        End RaiseEvent
    End Event

    Public Sub Dispatch(m As String)
        RaiseEvent Message(m)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New CustomPublisher()
        Dim h As Action(Of String) = Sub(m) Console.WriteLine("Got: " & m)
        AddHandler p.Message, h
        p.Dispatch("Hello")
        RemoveHandler p.Message, h
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Custom AddHandler", "Got: Hello", "Custom RemoveHandler"]
    );
}

#[test]
fn test_vb_event_generic_event_args() {
    let src = r#"
Imports System

Class DataEventArgs(Of T)
    Inherits EventArgs
    Public ReadOnly Property Data As T
    Public Sub New(d As T)
        Data = d
    End Sub
End Class

Class DataBroadcaster(Of T)
    Public Event DataReceived As EventHandler(Of DataEventArgs(Of T))
    Public Sub Broadcast(d As T)
        RaiseEvent DataReceived(Me, New DataEventArgs(Of T)(d))
    End Sub
End Class

Module Program
    Sub Main()
        Dim b As New DataBroadcaster(Of String)()
        AddHandler b.DataReceived, Sub(s, e) Console.WriteLine("Recv: " & e.Data)
        b.Broadcast("PayloadString")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Recv: PayloadString"]);
}

#[test]
fn test_vb_event_static_shared_event() {
    let src = r#"
Imports System

Class GlobalBus
    Public Shared Event GlobalMessage As Action(Of String)
    Public Shared Sub Broadcast(msg As String)
        RaiseEvent GlobalMessage(msg)
    End Sub
End Class

Module Program
    Sub Main()
        AddHandler GlobalBus.GlobalMessage, Sub(m) Console.WriteLine("Global: " & m)
        GlobalBus.Broadcast("Ping")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Global: Ping"]);
}

#[test]
fn test_vb_event_raising_event_inside_handler_reentrancy() {
    let src = r#"
Imports System

Class ChainEmitter
    Public Event Step1 As Action
    Public Event Step2 As Action

    Public Sub Run()
        RaiseEvent Step1()
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New ChainEmitter()
        AddHandler c.Step1, Sub()
            Console.WriteLine("Step1 Triggered")
            RaiseEvent c.Step2()
        End Sub
        AddHandler c.Step2, Sub() Console.WriteLine("Step2 Triggered")
        c.Run()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Step1 Triggered", "Step2 Triggered"]);
}

#[test]
fn test_vb_event_in_interface_raise_via_method() {
    let src = r#"
Imports System

Interface IClickable
    Event Click As EventHandler
    Sub DoClick()
End Interface

Class LinkLabel
    Implements IClickable
    Public Event Click As EventHandler Implements IClickable.Click
    Public Sub DoClick() Implements IClickable.DoClick
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As IClickable = New LinkLabel()
        AddHandler c.Click, Sub(s, e) Console.WriteLine("Link Clicked")
        c.DoClick()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Link Clicked"]);
}

#[test]
fn test_vb_event_custom_args_mutation_in_handler() {
    let src = r#"
Imports System

Class CancelEventArgs
    Inherits EventArgs
    Public Property Cancel As Boolean = False
End Class

Class Worker
    Public Event QueryCancel As EventHandler(Of CancelEventArgs)
    Public Function TryPerformWork() As Boolean
        Dim args As New CancelEventArgs()
        RaiseEvent QueryCancel(Me, args)
        Return Not args.Cancel
    End Function
End Class

Module Program
    Sub Main()
        Dim w As New Worker()
        AddHandler w.QueryCancel, Sub(sender, e) e.Cancel = True
        Console.WriteLine("Can Work: " & w.TryPerformWork())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Can Work: False"]);
}

#[test]
fn test_vb_event_handler_exception_handling() {
    let src = r#"
Imports System

Class FaultyEmitter
    Public Event Notify As Action
    Public Sub Fire()
        Try
            RaiseEvent Notify()
        Catch ex As Exception
            Console.WriteLine("Caught Exception: " & ex.Message)
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Dim fe As New FaultyEmitter()
        AddHandler fe.Notify, Sub() Throw New InvalidOperationException("Handler Failed")
        fe.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Caught Exception: Handler Failed"]);
}

#[test]
fn test_vb_event_raise_from_derived_class() {
    let src = r#"
Imports System

Class BaseNotifier
    Public Event Notice As Action
    Protected Sub RaiseNotice()
        RaiseEvent Notice()
    End Sub
End Class

Class DerivedNotifier
    Inherits BaseNotifier
    Public Sub TriggerNotice()
        RaiseNotice()
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New DerivedNotifier()
        AddHandler d.Notice, Sub() Console.WriteLine("Notice Triggered")
        d.TriggerNotice()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Notice Triggered"]);
}

#[test]
fn test_vb_event_multiple_handlers_same_method() {
    let src = r#"
Imports System

Class Counter
    Public Event Increment As Action
    Public Sub Count()
        RaiseEvent Increment()
    End Sub
End Class

Module Program
    Private total As Integer = 0
    Private Sub AddOne()
        total += 1
    End Sub

    Sub Main()
        Dim c As New Counter()
        AddHandler c.Increment, AddressOf AddOne
        AddHandler c.Increment, AddressOf AddOne
        c.Count()
        Console.WriteLine(total)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_event_get_invocation_list_via_reflection() {
    let src = r#"
Imports System
Imports System.Reflection

Class Subject
    Public Event Update As EventHandler
    Public Function GetSubscriberCount() As Integer
        Dim field As FieldInfo = GetType(Subject).GetField("UpdateEvent", BindingFlags.NonPublic Or BindingFlags.Instance)
        If field IsNot Nothing Then
            Dim del = TryCast(field.GetValue(Me), [Delegate])
            If del IsNot Nothing Then Return del.GetInvocationList().Length
        End If
        Return 0
    End Function
End Class

Module Program
    Sub Main()
        Dim s As New Subject()
        AddHandler s.Update, Sub(sender, args) End Sub
        AddHandler s.Update, Sub(sender, args) End Sub
        Console.WriteLine(s.GetSubscriberCount())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_event_struct_event_handler() {
    let src = r#"
Imports System

Structure ValuePublisher
    Public Event ValueChanged As Action(Of Integer)
    Public Sub Publish(val As Integer)
        RaiseEvent ValueChanged(val)
    End Sub
End Structure

Module Program
    Sub Main()
        Dim vp As New ValuePublisher()
        AddHandler vp.ValueChanged, Sub(v) Console.WriteLine("Val: " & v)
        vp.Publish(99)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Val: 99"]);
}

#[test]
fn test_vb_event_property_changed_interface_implementation() {
    let src = r#"
Imports System.ComponentModel

Class Person
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private _name As String
    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            If _name <> value Then
                _name = value
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Name"))
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim p As New Person()
        AddHandler p.PropertyChanged, Sub(s, e) Console.WriteLine("Changed: " & e.PropertyName)
        p.Name = "Bob"
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Changed: Name"]);
}

#[test]
fn test_vb_event_lambda_with_closure_state() {
    let src = r#"
Imports System

Class CounterEmitter
    Public Event Counted As Action(Of Integer)
    Public Sub Tick()
        For i As Integer = 1 To 3
            RaiseEvent Counted(i)
        Next
    End Sub
End Class

Module Program
    Sub Main()
        Dim sum As Integer = 0
        Dim emitter As New CounterEmitter()
        AddHandler emitter.Counted, Sub(val) sum += val
        emitter.Tick()
        Console.WriteLine("Sum: " & sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Sum: 6"]);
}

#[test]
fn test_vb_event_delegate_target_and_method_inspection() {
    let src = r#"
Imports System

Class TargetReceiver
    Public Sub OnEvent()
        Console.WriteLine("Receiver Action")
    End Sub
End Class

Class EventSource
    Public Event Action As Action
    Public Sub Fire()
        RaiseEvent Action()
    End Sub
End Class

Module Program
    Sub Main()
        Dim tr As New TargetReceiver()
        Dim es As New EventSource()
        AddHandler es.Action, AddressOf tr.OnEvent
        es.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Receiver Action"]);
}

#[test]
fn test_vb_event_custom_event_thread_sync_lock() {
    let src = r#"
Imports System

Class ThreadSafeEventPublisher
    Private syncObj As New Object()
    Private internalDelegate As Action

    Public Custom Event SecureEvent As Action
        AddHandler(value As Action)
            SyncLock syncObj
                internalDelegate = CType([Delegate].Combine(internalDelegate, value), Action)
            End SyncLock
        End AddHandler
        RemoveHandler(value As Action)
            SyncLock syncObj
                internalDelegate = CType([Delegate].Remove(internalDelegate, value), Action)
            End SyncLock
        End RemoveHandler
        RaiseEvent()
            Dim copy As Action = Nothing
            SyncLock syncObj
                copy = internalDelegate
            End SyncLock
            If copy IsNot Nothing Then copy()
        End RaiseEvent
    End Event

    Public Sub Run()
        RaiseEvent SecureEvent()
    End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New ThreadSafeEventPublisher()
        AddHandler pub.SecureEvent, Sub() Console.WriteLine("ThreadSafe Event Fired")
        pub.Run()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ThreadSafe Event Fired"]);
}

#[test]
fn test_vb_event_raising_with_null_arguments_allowed() {
    let src = r#"
Imports System

Class NullArgsEmitter
    Public Event CustomNotify As EventHandler
    Public Sub RaiseNullArgs()
        RaiseEvent CustomNotify(Nothing, Nothing)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New NullArgsEmitter()
        AddHandler e.CustomNotify, Sub(sender, args)
            Console.WriteLine("SenderIsNull=" & (sender Is Nothing) & "|ArgsIsNull=" & (args Is Nothing))
        End Sub
        e.RaiseNullArgs()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SenderIsNull=True|ArgsIsNull=True"]);
}
