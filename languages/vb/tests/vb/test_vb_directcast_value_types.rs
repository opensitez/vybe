use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DirectCast Operator Semantics & Strict Casting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_directcast_inheritance_hierarchy_succeeds() {
    let src = r#"
Class Base
End Class

Class Child
    Inherits Base
    Public ReadOnly Tag As String = "ChildTag"
End Class

Module Program
    Sub Main()
        Dim b As Base = New Child()
        Dim c As Child = DirectCast(b, Child)
        Console.WriteLine(c.Tag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ChildTag"]);
}

#[test]
fn test_vb_directcast_invalid_type_throws_invalid_cast_exception() {
    let src = r#"
Imports System

Class A
End Class

Class B
End Class

Module Program
    Sub Main()
        Dim obj As Object = New A()
        Try
            Dim b As B = DirectCast(obj, B)
        Catch ex As InvalidCastException
            Console.WriteLine("InvalidCastException Caught on DirectCast")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidCastException Caught on DirectCast"]
    );
}

#[test]
fn test_vb_directcast_boxed_value_type_requires_exact_type() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim boxedInt As Object = 42
        ' DirectCast to exact boxed type (Integer) succeeds!
        Dim n As Integer = DirectCast(boxedInt, Integer)
        Console.WriteLine(n)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_directcast_boxed_value_type_widening_throws_invalid_cast() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim boxedInt As Object = 42
        Try
            ' DirectCast to Double from boxed Integer throws InvalidCastException (unlike CType)!
            Dim d As Double = DirectCast(boxedInt, Double)
        Catch ex As InvalidCastException
            Console.WriteLine("InvalidCastException Caught on DirectCast Widening")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidCastException Caught on DirectCast Widening"]
    );
}

#[test]
fn test_vb_directcast_interface_cast_succeeds() {
    let src = r#"
Imports System

Interface IRunner
    Sub Run()
End Interface

Class Worker
    Implements IRunner
    Public Sub Run() Implements IRunner.Run
        Console.WriteLine("Worker Running")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As Object = New Worker()
        Dim runner As IRunner = DirectCast(obj, IRunner)
        runner.Run()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Worker Running"]);
}

#[test]
fn test_vb_directcast_null_reference_to_reference_type_returns_nothing() {
    let src = r#"
Class Item
End Class

Module Program
    Sub Main()
        Dim obj As Object = Nothing
        Dim i As Item = DirectCast(obj, Item)
        Console.WriteLine(i Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_directcast_null_reference_to_value_type_returns_default() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj As Object = Nothing
        Dim num As Integer = DirectCast(obj, Integer)
        Console.WriteLine(num)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_directcast_string_from_object() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj As Object = "DirectCastString"
        Dim s As String = DirectCast(obj, String)
        Console.WriteLine(s)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["DirectCastString"]);
}

#[test]
fn test_vb_directcast_enum_exact_underlying_type_cast() {
    let src = r#"
Enum Color
    Red = 1
End Enum

Module Program
    Sub Main()
        Dim obj As Object = Color.Red
        Dim c As Color = DirectCast(obj, Color)
        Console.WriteLine(c.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Red"]);
}

#[test]
fn test_vb_directcast_boxed_struct_exact_match() {
    let src = r#"
Structure Point2D
    Public X, Y As Integer
End Structure

Module Program
    Sub Main()
        Dim p As New Point2D With {.X = 10, .Y = 20}
        Dim boxed As Object = p
        Dim restored As Point2D = DirectCast(boxed, Point2D)
        Console.WriteLine(restored.X & "," & restored.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_directcast_does_not_invoke_custom_conversion_operator() {
    let src = r#"
Imports System

Class Money
    Public Amount As Decimal
    Public Sub New(a As Decimal)
        Amount = a
    End Sub

    Public Shared Widening Operator CType(a As Decimal) As Money
        Return New Money(a)
    End Shared Widening Operator
End Class

Module Program
    Sub Main()
        Dim boxed As Object = 99.9D
        Try
            ' DirectCast does not call user-defined CType conversion operator!
            Dim m As Money = DirectCast(boxed, Money)
        Catch ex As InvalidCastException
            Console.WriteLine("InvalidCastException Caught on Custom Operator DirectCast")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidCastException Caught on Custom Operator DirectCast"]
    );
}

#[test]
fn test_vb_directcast_generic_class_strict_type_checking() {
    let src = r#"
Class Container(Of T)
    Public Element As T
End Class

Module Program
    Sub Main()
        Dim c As Object = New Container(Of String) With {.Element = "Inside"}
        Dim typed As Container(Of String) = DirectCast(c, Container(Of String))
        Console.WriteLine(typed.Element)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Inside"]);
}

#[test]
fn test_vb_directcast_array_exact_type_match() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30}
        Dim obj As Object = numbers
        Dim arr As Integer() = DirectCast(obj, Integer())
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_directcast_array_element_type_widening_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Byte() = {1, 2, 3}
        Dim obj As Object = bytes
        Try
            ' Byte() cannot be DirectCast to Integer()!
            Dim ints As Integer() = DirectCast(obj, Integer())
        Catch ex As InvalidCastException
            Console.WriteLine("InvalidCastException Caught on Array Type Mismatch")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidCastException Caught on Array Type Mismatch"]
    );
}

#[test]
fn test_vb_directcast_delegate_to_exact_delegate_type() {
    let src = r#"
Imports System

Delegate Function CustomFunc(x As Integer) As Integer

Module Program
    Private Function DoubleIt(x As Integer) As Integer
        Return x * 2
    End Function

    Sub Main()
        Dim del As Object = New CustomFunc(AddressOf DoubleIt)
        Dim cf As CustomFunc = DirectCast(del, CustomFunc)
        Console.WriteLine(cf(15))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30"]);
}

#[test]
fn test_vb_directcast_nullable_type_from_boxed_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim boxedInt As Object = 100
        Dim n As Integer? = DirectCast(boxedInt, Integer?)
        Console.WriteLine(n.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_directcast_value_tuple_exact_generic_arguments() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim tupleObj As Object = ValueTuple.Create(1, "A")
        Dim t = DirectCast(tupleObj, ValueTuple(Of Integer, String))
        Console.WriteLine(t.Item1 & ":" & t.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:A"]);
}

#[test]
fn test_vb_directcast_multidimensional_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim grid(,) As Integer = {{1, 2}, {3, 4}}
        Dim obj As Object = grid
        Dim restored(,) As Integer = DirectCast(obj, Integer(,))
        Console.WriteLine(restored(1, 1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}

#[test]
fn test_vb_directcast_derived_class_to_base_class() {
    let src = r#"
Class Parent
    Public Value As Integer = 50
End Class

Class SubChild
    Inherits Parent
End Class

Module Program
    Sub Main()
        Dim child As SubChild = New SubChild()
        Dim p As Parent = DirectCast(child, Parent)
        Console.WriteLine(p.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_directcast_exception_hierarchy_exact_match() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim exObj As Object = New FormatException("Invalid Format")
        Dim fe As FormatException = DirectCast(exObj, FormatException)
        Console.WriteLine(fe.Message)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Invalid Format"]);
}
