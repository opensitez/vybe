use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: NotInheritable, NotOverridable & Sealed Modifiers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_notinheritable_class_instantiation() {
    let src = r#"
NotInheritable Class SealedClass
    Public Function GetValue() As String
        Return "SealedValue"
    End Function
End Class

Module Program
    Sub Main()
        Dim obj As New SealedClass()
        Console.WriteLine(obj.GetValue())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SealedValue"]);
}

#[test]
fn test_vb_notoverridable_method_override_seal() {
    let src = r#"
Class BaseClass
    Public Overridable Sub Display()
        Console.WriteLine("Base Display")
    End Sub
End Class

Class MiddleClass
    Inherits BaseClass
    Public NotOverridable Overrides Sub Display()
        Console.WriteLine("Middle Sealed Display")
    End Sub
End Class

Module Program
    Sub Main()
        Dim m As BaseClass = New MiddleClass()
        m.Display()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Middle Sealed Display"]);
}

#[test]
fn test_vb_notinheritable_class_with_static_shared_members() {
    let src = r#"
NotInheritable Class UtilityHelper
    Public Shared Function Multiply(a As Integer, b As Integer) As Integer
        Return a * b
    End Function
End Class

Module Program
    Sub Main()
        Console.WriteLine(UtilityHelper.Multiply(6, 7))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_notinheritable_class_implementing_interface() {
    let src = r#"
Interface IService
    Sub Serve()
End Interface

NotInheritable Class ConcreteService
    Implements IService
    Public Sub Serve() Implements IService.Serve
        Console.WriteLine("Serving Concrete")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As IService = New ConcreteService()
        s.Serve()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Serving Concrete"]);
}

#[test]
fn test_vb_notinheritable_class_inheriting_from_abstract_class() {
    let src = r#"
MustInherit Class AbstractWorker
    Public MustOverride Sub Work()
End Class

NotInheritable Class FinalWorker
    Inherits AbstractWorker
    Public Overrides Sub Work()
        Console.WriteLine("Final Work Completed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim w As AbstractWorker = New FinalWorker()
        w.Work()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Final Work Completed"]);
}

#[test]
fn test_vb_notoverridable_property() {
    let src = r#"
Class BaseComponent
    Public Overridable Property Title As String = "Base"
End Class

Class FixedComponent
    Inherits BaseComponent
    Public NotOverridable Overrides Property Title As String
        Get
            Return "FixedTitle"
        End Get
        Set(value As String)
            ' Ignore
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim c As BaseComponent = New FixedComponent()
        Console.WriteLine(c.Title)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FixedTitle"]);
}

#[test]
fn test_vb_notinheritable_generic_class() {
    let src = r#"
NotInheritable Class Container(Of T)
    Public Value As T
    Public Sub New(v As T)
        Value = v
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Container(Of Integer)(100)
        Console.WriteLine(c.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_notoverridable_event_handler() {
    let src = r#"
Imports System

Class BaseEmitter
    Public Overridable Custom Event Action As EventHandler
        AddHandler(value As EventHandler)
        End AddHandler
        RemoveHandler(value As EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
        End RaiseEvent
    End Event
End Class

Class FixedEmitter
    Inherits BaseEmitter
    Public NotOverridable Overrides Custom Event Action As EventHandler
        AddHandler(value As EventHandler)
            Console.WriteLine("Handler Added to Fixed")
        End AddHandler
        RemoveHandler(value As EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
        End RaiseEvent
    End Event
End Class

Module Program
    Sub Main()
        Dim e As BaseEmitter = New FixedEmitter()
        AddHandler e.Action, Sub(sender, args) End Sub
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Handler Added to Fixed"]);
}

#[test]
fn test_vb_notinheritable_class_private_constructor_pattern() {
    let src = r#"
NotInheritable Class Singleton
    Public Shared ReadOnly Instance As New Singleton()
    Private Sub New()
    End Sub
    Public Function GetName() As String
        Return "SingletonInstance"
    End Function
End Class

Module Program
    Sub Main()
        Console.WriteLine(Singleton.Instance.GetName())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SingletonInstance"]);
}

#[test]
fn test_vb_notinheritable_struct_inherent_behavior() {
    let src = r#"
Structure FixedPoint
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim p As New FixedPoint(5, 10)
        Console.WriteLine(p.X & "," & p.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5,10"]);
}

#[test]
fn test_vb_notoverridable_method_calling_base_method() {
    let src = r#"
Class Parent
    Public Overridable Function Log(msg As String) As String
        Return "Parent: " & msg
    End Function
End Class

Class Child
    Inherits Parent
    Public NotOverridable Overrides Function Log(msg As String) As String
        Return MyBase.Log(msg) & " (Child Sealed)"
    End Function
End Class

Module Program
    Sub Main()
        Dim p As Parent = New Child()
        Console.WriteLine(p.Log("Message"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Parent: Message (Child Sealed)"]);
}

#[test]
fn test_vb_notinheritable_enum_type_inherent_behavior() {
    let src = r#"
Enum FinalStatus
    Success = 1
    Failure = 2
End Enum

Module Program
    Sub Main()
        Dim s As FinalStatus = FinalStatus.Success
        Console.WriteLine(s.ToString() & "=" & CInt(s))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Success=1"]);
}

#[test]
fn test_vb_notoverridable_shadows_combination() {
    let src = r#"
Class GrandParent
    Public Overridable Sub Show() : Console.WriteLine("GrandParent") : End Sub
End Class

Class Parent
    Inherits GrandParent
    Public Overrides Sub Show() : Console.WriteLine("Parent") : End Sub
End Class

Class Child
    Inherits Parent
    Public Shadows Sub Show() : Console.WriteLine("Child Shadow") : End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Child()
        c.Show()
        Dim p As Parent = c
        p.Show()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Child Shadow", "Parent"]);
}

#[test]
fn test_vb_notinheritable_attribute_usage() {
    let src = r#"
Imports System

<AttributeUsage(AttributeTargets.Class)>
NotInheritable Class CustomInfoAttribute
    Inherits Attribute
    Public Property Description As String
    Public Sub New(desc As String)
        Description = desc
    End Sub
End Class

<CustomInfo("TestClass")>
Class Target
End Class

Module Program
    Sub Main()
        Dim attrs = GetType(Target).GetCustomAttributes(GetType(CustomInfoAttribute), False)
        Dim info = CType(attrs(0), CustomInfoAttribute)
        Console.WriteLine(info.Description)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TestClass"]);
}

#[test]
fn test_vb_notinheritable_nested_class() {
    let src = r#"
Class Outer
    NotInheritable Class Inner
        Public Shared Function InnerMethod() As String
            Return "InnerResult"
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Console.WriteLine(Outer.Inner.InnerMethod())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["InnerResult"]);
}

#[test]
fn test_vb_notoverridable_indexer_property() {
    let src = r#"
Class BaseContainer
    Default Public Overridable Property Item(index As Integer) As String
        Get
            Return "Base_" & index
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Class FixedContainer
    Inherits BaseContainer
    Default Public NotOverridable Overrides Property Item(index As Integer) As String
        Get
            Return "Fixed_" & index
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim b As BaseContainer = New FixedContainer()
        Console.WriteLine(b(5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Fixed_5"]);
}

#[test]
fn test_vb_notinheritable_extension_method_container() {
    let src = r#"
Imports System.Runtime.CompilerServices

<Extension()>
Module StringExtensions
    <Extension()>
    Public Function Exclaim(s As String) As String
        Return s & "!"
    End Function
End Module

Module Program
    Sub Main()
        Dim text As String = "Hello"
        Console.WriteLine(text.Exclaim())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello!"]);
}

#[test]
fn test_vb_notoverridable_multiple_overloads() {
    let src = r#"
Class BaseCalc
    Public Overridable Function Compute(x As Integer) As Integer
        Return x * 2
    End Function
    Public Overridable Function Compute(x As Double) As Double
        Return x * 2.0
    End Function
End Class

Class SealedCalc
    Inherits BaseCalc
    Public NotOverridable Overrides Function Compute(x As Integer) As Integer
        Return x * 3
    End Function
End Class

Module Program
    Sub Main()
        Dim b As BaseCalc = New SealedCalc()
        Console.WriteLine(b.Compute(10) & "|" & b.Compute(10.0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30|20"]);
}

#[test]
fn test_vb_notinheritable_record_class_simulation() {
    let src = r#"
NotInheritable Class ImmutableData
    Public ReadOnly Property ID As Integer
    Public ReadOnly Property Value As String
    Public Sub New(id As Integer, val As String)
        Me.ID = id : Me.Value = val
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New ImmutableData(1, "Val")
        Console.WriteLine(d.ID & ":" & d.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:Val"]);
}

#[test]
fn test_vb_notinheritable_type_typeof_checking() {
    let src = r#"
NotInheritable Class SealedType
End Class

Module Program
    Sub Main()
        Dim t = GetType(SealedType)
        Console.WriteLine(t.IsSealed)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
