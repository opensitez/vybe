use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: TryCast Operator Semantics & Safe Casting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_trycast_compatible_reference_type_succeeds() {
    let src = r#"
Class Animal
End Class

Class Dog
    Inherits Animal
    Public ReadOnly Name As String = "Rover"
End Class

Module Program
    Sub Main()
        Dim a As Animal = New Dog()
        Dim d As Dog = TryCast(a, Dog)
        Console.WriteLine(d IsNot Nothing AndAlso d.Name = "Rover")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_trycast_incompatible_reference_type_returns_nothing() {
    let src = r#"
Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Class Cat
    Inherits Animal
End Class

Module Program
    Sub Main()
        Dim a As Animal = New Dog()
        Dim c As Cat = TryCast(a, Cat)
        Console.WriteLine(c Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_trycast_interface_implementation_succeeds() {
    let src = r#"
Imports System

Interface IPlayable
    Sub Play()
End Interface

Class Widget
    Implements IPlayable
    Public Sub Play() Implements IPlayable.Play
        Console.WriteLine("Widget Playing")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As Object = New Widget()
        Dim p As IPlayable = TryCast(obj, IPlayable)
        Console.WriteLine(p IsNot Nothing)
        If p IsNot Nothing Then p.Play()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "Widget Playing"]);
}

#[test]
fn test_vb_trycast_null_source_returns_nothing() {
    let src = r#"
Class Shape
End Class

Module Program
    Sub Main()
        Dim obj As Object = Nothing
        Dim s As Shape = TryCast(obj, Shape)
        Console.WriteLine(s Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_trycast_generic_class_inheritance() {
    let src = r#"
Class BaseContainer(Of T)
End Class

Class DerivedContainer(Of T)
    Inherits BaseContainer(Of T)
End Class

Module Program
    Sub Main()
        Dim baseObj As BaseContainer(Of Integer) = New DerivedContainer(Of Integer)()
        Dim derivedObj As DerivedContainer(Of Integer) = TryCast(baseObj, DerivedContainer(Of Integer))
        Console.WriteLine(derivedObj IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_trycast_string_from_object() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj As Object = "Hello World"
        Dim str As String = TryCast(obj, String)
        Console.WriteLine(str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_trycast_array_reference_type_cast() {
    let src = r#"
Module Program
    Sub Main()
        Dim strings As String() = {"A", "B"}
        Dim objs As Object() = strings ' Array covariance
        Dim strArr As String() = TryCast(objs, String())
        Console.WriteLine(strArr IsNot Nothing AndAlso strArr.Length = 2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_trycast_multicast_delegate_conversion() {
    let src = r#"
Imports System

Delegate Sub MultiHandler()

Module Program
    Private Sub Target()
        Console.WriteLine("Target Called")
    End Sub

    Sub Main()
        Dim del As [Delegate] = New MultiHandler(AddressOf Target)
        Dim typedDel As MultiHandler = TryCast(del, MultiHandler)
        Console.WriteLine(typedDel IsNot Nothing)
        If typedDel IsNot Nothing Then typedDel()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "Target Called"]);
}

#[test]
fn test_vb_trycast_nullable_type_target_supported() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' TryCast can cast boxed value type to Nullable(Of T)
        Dim boxed As Object = 42
        Dim n As Integer? = TryCast(boxed, Integer?)
        Console.WriteLine(n.HasValue & "|" & n.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|42"]);
}

#[test]
fn test_vb_trycast_incompatible_boxed_value_type_returns_null() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' Boxed Double cast to Nullable(Of Integer) via TryCast returns Nothing!
        Dim boxed As Object = 3.14
        Dim n As Integer? = TryCast(boxed, Integer?)
        Console.WriteLine(n.HasValue)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_trycast_does_not_invoke_custom_conversion_operators() {
    let src = r#"
Class ComplexNum
    Public Real As Double
    Public Sub New(r As Double)
        Real = r
    End Sub

    Public Shared Narrowing Operator CType(d As Double) As ComplexNum
        Return New ComplexNum(d)
    End Narrowing Operator
End Class

Module Program
    Sub Main()
        Dim dblObj As Object = 10.5
        ' TryCast only checks inheritance/interface, does not call CType operator!
        Dim cn As ComplexNum = TryCast(dblObj, ComplexNum)
        Console.WriteLine(cn Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_trycast_unrelated_interfaces_returns_nothing() {
    let src = r#"
Interface IFoo
End Interface

Interface IBar
End Interface

Class ImplementFoo
    Implements IFoo
End Class

Module Program
    Sub Main()
        Dim foo As IFoo = New ImplementFoo()
        Dim bar As IBar = TryCast(foo, IBar)
        Console.WriteLine(bar Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_trycast_exception_class_hierarchy() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ex As Exception = New InvalidOperationException("Failed")
        Dim ioe As InvalidOperationException = TryCast(ex, InvalidOperationException)
        Dim ae As ArgumentException = TryCast(ex, ArgumentException)
        Console.WriteLine((ioe IsNot Nothing) & "|" & (ae Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_trycast_in_if_condition_assignment() {
    let src = r#"
Class Account
    Public Property Id As Integer = 101
End Class

Module Program
    Sub Main()
        Dim obj As Object = New Account()
        Dim acc As Account = TryCast(obj, Account)
        If acc IsNot Nothing Then
            Console.WriteLine("Account ID: " & acc.Id)
        End If
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Account ID: 101"]);
}

#[test]
fn test_vb_trycast_generic_method_type_parameter() {
    let src = r#"
Module Program
    Private Function SafeCast(Of T As Class)(input As Object) As T
        Return TryCast(input, T)
    End Function

    Sub Main()
        Dim strObj As Object = "GenericCast"
        Dim intObj As Object = 100
        Dim resStr = SafeCast(Of String)(strObj)
        Dim resIntStr = SafeCast(Of String)(intObj)
        Console.WriteLine(resStr & "|" & (resIntStr Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["GenericCast|True"]);
}

#[test]
fn test_vb_trycast_enum_boxed_to_nullable_enum() {
    let src = r#"
Imports System

Enum Status
    Active
End Enum

Module Program
    Sub Main()
        Dim boxed As Object = Status.Active
        Dim s As Status? = TryCast(boxed, Status?)
        Console.WriteLine(s.HasValue & "|" & s.Value.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|Active"]);
}

#[test]
fn test_vb_trycast_derived_interface_inheritance() {
    let src = r#"
Interface IBase
End Interface

Interface IDerived
    Inherits IBase
End Interface

Class Implementation
    Implements IDerived
End Class

Module Program
    Sub Main()
        Dim impl As Object = New Implementation()
        Dim b As IBase = TryCast(impl, IBase)
        Console.WriteLine(b IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_trycast_array_incompatible_element_type() {
    let src = r#"
Module Program
    Sub Main()
        Dim ints As Object() = New Object() {1, 2, 3}
        Dim strings As String() = TryCast(ints, String())
        Console.WriteLine(strings Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_trycast_value_tuple_boxed_reference() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim tupleBox As Object = ValueTuple.Create(10, "A")
        Dim t As ValueTuple(Of Integer, String)? = TryCast(tupleBox, ValueTuple(Of Integer, String)?)
        Console.WriteLine(t.HasValue & "|" & t.Value.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|A"]);
}

#[test]
fn test_vb_trycast_same_type_identity_cast() {
    let src = r#"
Class Node
End Class

Module Program
    Sub Main()
        Dim n As New Node()
        Dim n2 As Node = TryCast(n, Node)
        Console.WriteLine(Object.ReferenceEquals(n, n2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
