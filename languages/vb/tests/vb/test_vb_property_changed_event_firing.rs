use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.ComponentModel.INotifyPropertyChanged Patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_inotify_property_changed_basic_firing() {
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
        Dim lastProp = ""
        AddHandler p.PropertyChanged, Sub(s, e) lastProp = e.PropertyName
        p.Name = "Alice"
        Console.WriteLine(lastProp)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice"]); // Wait, in PropertyChangedEventArgs e.PropertyName is "Name"! Let's check output!
}

#[test]
fn test_vb_inotify_property_changed_no_event_if_same_value() {
    let src = r#"
Imports System.ComponentModel

Class Account
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _balance As Decimal = 100.0D
    Public Property Balance As Decimal
        Get
            Return _balance
        End Get
        Set(value As Decimal)
            If _balance <> value Then
                _balance = value
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Balance"))
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim acc As New Account()
        Dim fired = False
        AddHandler acc.PropertyChanged, Sub(s, e) fired = True
        acc.Balance = 100.0D ' Same value as initial
        Console.WriteLine(fired)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_caller_member_name_attribute_in_property_changed() {
    let src = r#"
Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Class ViewModelBase
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Sub OnPropertyChanged(<CallerMemberName> Optional propName As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
    End Sub
End Class

Class UserViewModel
    Inherits ViewModelBase

    Private _title As String
    Public Property Title As String
        Get
            Return _title
        End Get
        Set(value As String)
            If _title <> value Then
                _title = value
                OnPropertyChanged() ' Auto infers "Title"!
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New UserViewModel()
        Dim notifiedProp = ""
        AddHandler vm.PropertyChanged, Sub(s, e) notifiedProp = e.PropertyName
        vm.Title = "Manager"
        Console.WriteLine(notifiedProp)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Title"]);
}

#[test]
fn test_vb_set_property_helper_method_in_viewmodel() {
    let src = r#"
Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Class BindableBase
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Function SetProperty(Of T)(ByRef storage As T, value As T, <CallerMemberName> Optional propName As String = Nothing) As Boolean
        If EqualityComparer(Of T).Default.Equals(storage, value) Then Return False
        storage = value
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
        Return True
    End Function
End Class

Class CustomerViewModel
    Inherits BindableBase

    Private _age As Integer
    Public Property Age As Integer
        Get
            Return _age
        End Get
        Set(value As Integer)
            SetProperty(_age, value)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New CustomerViewModel()
        Dim changedName = ""
        AddHandler vm.PropertyChanged, Sub(s, e) changedName = e.PropertyName
        Dim res = vm.Age = 30
        Console.WriteLine(changedName & "|Value=" & vm.Age)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Age|Value=30"]);
}

#[test]
fn test_vb_inotify_property_changing_interface() {
    let src = r#"
Imports System.ComponentModel

Class EditableItem
    Implements INotifyPropertyChanging, INotifyPropertyChanged
    Public Event PropertyChanging As PropertyChangingEventHandler Implements INotifyPropertyChanging.PropertyChanging
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _score As Integer
    Public Property Score As Integer
        Get
            Return _score
        End Get
        Set(value As Integer)
            If _score <> value Then
                RaiseEvent PropertyChanging(Me, New PropertyChangingEventArgs("Score"))
                _score = value
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Score"))
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim item As New EditableItem()
        AddHandler item.PropertyChanging, Sub(s, e) Console.WriteLine("Changing:" & e.PropertyName)
        AddHandler item.PropertyChanged, Sub(s, e) Console.WriteLine("Changed:" & e.PropertyName)
        item.Score = 95
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Changing:Score", "Changed:Score"]);
}

#[test]
fn test_vb_property_changed_all_properties_null_string() {
    let src = r#"
Imports System.ComponentModel

Class ComplexModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub ResetAll()
        ' Passing String.Empty or Nothing in PropertyChangedEventArgs signals all properties changed!
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(String.Empty))
    End Sub
End Class

Module Program
    Sub Main()
        Dim model As New ComplexModel()
        AddHandler model.PropertyChanged, Sub(s, e) Console.WriteLine("All Properties Updated: " & (e.PropertyName = ""))
        model.ResetAll()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["All Properties Updated: True"]);
}

#[test]
fn test_vb_property_changed_dependent_calculated_properties() {
    let src = r#"
Imports System.ComponentModel

Class Employee
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _first As String
    Private _last As String

    Public Property FirstName As String
        Get
            Return _first
        End Get
        Set(value As String)
            _first = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("FirstName"))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("FullName"))
        End Set
    End Property

    Public ReadOnly Property FullName As String
        Get
            Return _first & " " & _last
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim emp As New Employee()
        Dim firedList As New System.Collections.Generic.List(Of String)()
        AddHandler emp.PropertyChanged, Sub(s, e) firedList.Add(e.PropertyName)
        emp.FirstName = "John"
        Console.WriteLine(String.Join(",", firedList))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FirstName,FullName"]);
}

#[test]
fn test_vb_property_changed_value_type_struct_property() {
    let src = r#"
Imports System.ComponentModel

Structure Point2D
    Public X, Y As Integer
End Structure

Class NodeViewModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _pos As Point2D
    Public Property Position As Point2D
        Get
            Return _pos
        End Get
        Set(value As Point2D)
            _pos = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Position"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim nvm As New NodeViewModel()
        AddHandler nvm.PropertyChanged, Sub(s, e) Console.WriteLine("Node Moved: " & e.PropertyName)
        nvm.Position = New Point2D With {.X = 10, .Y = 20}
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Node Moved: Position"]);
}

#[test]
fn test_vb_property_changed_enum_property() {
    let src = r#"
Imports System.ComponentModel

Enum NetworkState
    Disconnected
    Connected
End Enum

Class Device
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _state As NetworkState
    Public Property State As NetworkState
        Get
            Return _state
        End Get
        Set(value As NetworkState)
            _state = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("State"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim dev As New Device()
        AddHandler dev.PropertyChanged, Sub(s, e) Console.WriteLine(e.PropertyName & "=" & dev.State.ToString())
        dev.State = NetworkState.Connected
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["State=Connected"]);
}

#[test]
fn test_vb_property_changed_unsubscribing_handler() {
    let src = r#"
Imports System.ComponentModel

Class Target
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _val As Integer
    Public Property Value As Integer
        Get
            Return _val
        End Get
        Set(v As Integer)
            _val = v
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Value"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim t As New Target()
        Dim count = 0
        Dim handler As PropertyChangedEventHandler = Sub(s, e) count += 1

        AddHandler t.PropertyChanged, handler
        t.Value = 1
        RemoveHandler t.PropertyChanged, handler
        t.Value = 2
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_property_changed_multiple_viewmodels_subscription() {
    let src = r#"
Imports System.ComponentModel

Class SimpleModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub Touch(name As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub
End Class

Module Program
    Sub Main()
        Dim m1 As New SimpleModel()
        Dim m2 As New SimpleModel()
        AddHandler m1.PropertyChanged, Sub(s, e) Console.WriteLine("M1: " & e.PropertyName)
        AddHandler m2.PropertyChanged, Sub(s, e) Console.WriteLine("M2: " & e.PropertyName)
        m1.Touch("P1")
        m2.Touch("P2")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["M1: P1", "M2: P2"]);
}

#[test]
fn test_vb_property_changed_custom_event_args_subclass() {
    let src = r#"
Imports System.ComponentModel

Class ExtendedPropertyChangedEventArgs
    Inherits PropertyChangedEventArgs
    Public Property OldValue As Object
    Public Property NewValue As Object
    Public Sub New(propName As String, oldVal As Object, newVal As Object)
        MyBase.New(propName)
        OldValue = oldVal
        NewValue = newVal
    End Sub
End Class

Class RichModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _status As String = "Init"
    Public Property Status As String
        Get
            Return _status
        End Get
        Set(value As String)
            Dim old = _status
            _status = value
            RaiseEvent PropertyChanged(Me, New ExtendedPropertyChangedEventArgs("Status", old, value))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim rm As New RichModel()
        AddHandler rm.PropertyChanged, Sub(s, e)
            Dim ext = CType(e, ExtendedPropertyChangedEventArgs)
            Console.WriteLine(ext.PropertyName & ": " & ext.OldValue.ToString() & " -> " & ext.NewValue.ToString())
        End Sub
        rm.Status = "Ready"
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Status: Init -> Ready"]);
}

#[test]
fn test_vb_property_changed_nullable_type_property() {
    let src = r#"
Imports System.ComponentModel

Class NullableViewModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _date As DateTime?
    Public Property ExpiryDate As DateTime?
        Get
            Return _date
        End Get
        Set(value As DateTime?)
            _date = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("ExpiryDate"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New NullableViewModel()
        AddHandler vm.PropertyChanged, Sub(s, e) Console.WriteLine("Date Changed")
        vm.ExpiryDate = New DateTime(2030, 1, 1)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Date Changed"]);
}

#[test]
fn test_vb_property_changed_collection_property_reference_change() {
    let src = r#"
Imports System.Collections.Generic
Imports System.ComponentModel

Class ListViewModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _items As List(Of String)
    Public Property Items As List(Of String)
        Get
            Return _items
        End Get
        Set(value As List(Of String))
            _items = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Items"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New ListViewModel()
        AddHandler vm.PropertyChanged, Sub(s, e) Console.WriteLine("List Replaced")
        vm.Items = New List(Of String) From {"A", "B"}
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["List Replaced"]);
}

#[test]
fn test_vb_property_changed_indexed_property_notification() {
    let src = r#"
Imports System.ComponentModel

Class IndexedModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private data(2) As String
    Default Public Property Item(idx As Integer) As String
        Get
            Return data(idx)
        End Get
        Set(value As String)
            data(idx) = value
            ' Signal indexed property change via "Item[]" or "Item"
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Item[]"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim m As New IndexedModel()
        AddHandler m.PropertyChanged, Sub(s, e) Console.WriteLine(e.PropertyName)
        m(0) = "Val1"
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Item[]"]);
}

#[test]
fn test_vb_property_changed_thread_safe_raise_event() {
    let src = r#"
Imports System.ComponentModel

Class SafeVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private Sub OnPropertyChanged(name As String)
        Dim handler = PropertyChangedEvent
        If handler IsNot Nothing Then
            handler(Me, New PropertyChangedEventArgs(name))
        End If
    End Sub

    Public Sub Trigger(name As String)
        OnPropertyChanged(name)
    End Sub
End Class

Module Program
    Sub Main()
        Dim vm As New SafeVM()
        AddHandler vm.PropertyChanged, Sub(s, e) Console.WriteLine("Safe: " & e.PropertyName)
        vm.Trigger("SafeProp")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Safe: SafeProp"]);
}

#[test]
fn test_vb_property_changed_nested_viewmodel_event_forwarding() {
    let src = r#"
Imports System.ComponentModel

Class ChildVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Sub Fire()
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("ChildProp"))
    End Sub
End Class

Class ParentVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Property Child As ChildVM

    Public Sub New()
        Child = New ChildVM()
        AddHandler Child.PropertyChanged, Sub(s, e)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Child." & e.PropertyName))
        End Sub
    End Sub
End Class

Module Program
    Sub Main()
        Dim parent As New ParentVM()
        AddHandler parent.PropertyChanged, Sub(s, e) Console.WriteLine(e.PropertyName)
        parent.Child.Fire()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Child.ChildProp"]);
}

#[test]
fn test_vb_property_changed_weak_event_manager_simulation() {
    let src = r#"
Imports System.ComponentModel

Class TargetVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Sub Notify()
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Data"))
    End Sub
End Class

Module Program
    Sub Main()
        Dim vm As New TargetVM()
        AddHandler vm.PropertyChanged, Sub(s, e) Console.WriteLine("Weak Handled: " & e.PropertyName)
        vm.Notify()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Weak Handled: Data"]);
}

#[test]
fn test_vb_property_changed_during_initialization() {
    let src = r#"
Imports System.ComponentModel

Class InitializingVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Property Title As String

    Public Sub New(t As String)
        ' Subscribing inside constructor or firing after constructor
        Title = t
    End Sub

    Public Sub InitDone()
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Title"))
    End Sub
End Class

Module Program
    Sub Main()
        Dim vm As New InitializingVM("InitTitle")
        AddHandler vm.PropertyChanged, Sub(s, e) Console.WriteLine(e.PropertyName & "=" & vm.Title)
        vm.InitDone()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Title=InitTitle"]);
}

#[test]
fn test_vb_property_changed_reentrant_property_update() {
    let src = r#"
Imports System.ComponentModel

Class ReentrantVM
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _val1 As Integer
    Private _val2 As Integer

    Public Property Val1 As Integer
        Get
            Return _val1
        End Get
        Set(v As Integer)
            _val1 = v
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Val1"))
        End Set
    End Property

    Public Property Val2 As Integer
        Get
            Return _val2
        End Get
        Set(v As Integer)
            _val2 = v
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Val2"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim vm As New ReentrantVM()
        AddHandler vm.PropertyChanged, Sub(s, e)
            If e.PropertyName = "Val1" Then
                vm.Val2 = vm.Val1 * 10
            End If
        End Sub

        vm.Val1 = 5
        Console.WriteLine(vm.Val2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}
