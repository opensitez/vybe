use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Is & IsNot Operators, TypeOf...Is & Object Identity
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_is_operator_same_object_reference() {
    let src = r#"
Module Program
    Class Widget
    End Class

    Sub Main()
        Dim w1 As New Widget()
        Dim w2 As Widget = w1
        Console.WriteLine(w1 Is w2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_is_operator_different_object_references() {
    let src = r#"
Module Program
    Class Widget
    End Class

    Sub Main()
        Dim w1 As New Widget()
        Dim w2 As New Widget()
        Console.WriteLine(w1 Is w2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_is_operator_null_reference_check() {
    let src = r#"
Module Program
    Class Service
    End Class

    Sub Main()
        Dim s As Service = Nothing
        Console.WriteLine(s Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_isnot_operator_null_reference_check() {
    let src = r#"
Module Program
    Class Service
    End Class

    Sub Main()
        Dim s As New Service()
        Console.WriteLine(s IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_typeof_is_exact_type() {
    let src = r#"
Module Program
    Class Animal
    End Class

    Sub Main()
        Dim a As Object = New Animal()
        Console.WriteLine(TypeOf a Is Animal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_typeof_is_derived_class() {
    let src = r#"
Module Program
    Class Animal
    End Class

    Class Dog
        Inherits Animal
    End Class

    Sub Main()
        Dim d As Object = New Dog()
        Console.WriteLine(TypeOf d Is Animal & "|" & TypeOf d Is Dog)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_typeof_isnot_derived_class() {
    let src = r#"
Module Program
    Class Animal
    End Class

    Class Cat
        Inherits Animal
    End Class

    Sub Main()
        Dim a As Object = New Animal()
        Console.WriteLine(TypeOf a IsNot Cat)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_typeof_is_interface_implementation() {
    let src = r#"
Module Program
    Interface IRunnable
    End Interface

    Class TaskRunner
        Implements IRunnable
    End Class

    Sub Main()
        Dim obj As Object = New TaskRunner()
        Console.WriteLine(TypeOf obj Is IRunnable)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_typeof_is_value_type_boxed() {
    let src = r#"
Module Program
    Sub Main()
        Dim boxedInt As Object = 42
        Dim boxedDouble As Object = 3.14
        Console.WriteLine((TypeOf boxedInt Is Integer) & "|" & (TypeOf boxedDouble Is Double))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_typeof_is_null_object_returns_false() {
    let src = r#"
Module Program
    Class Person
    End Class

    Sub Main()
        Dim p As Object = Nothing
        Console.WriteLine(TypeOf p Is Person)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_is_operator_string_interning_reference_check() {
    let src = r#"
Module Program
    Sub Main()
        Dim s1 As String = "LiteralString"
        Dim s2 As String = "LiteralString"
        Console.WriteLine(s1 Is s2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_is_operator_string_non_interned_reference_check() {
    let src = r#"
Module Program
    Sub Main()
        Dim s1 As String = New String({"A"c, "B"c})
        Dim s2 As String = New String({"A"c, "B"c})
        Console.WriteLine((s1 = s2) & "|" & (s1 Is s2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_typeof_is_generic_type_instance() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As Object = New List(Of String)()
        Console.WriteLine(TypeOf list Is List(Of String) & "|" & TypeOf list IsIsNot List(Of Integer))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_typeof_is_nullable_value_type() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim n1 As Object = CType(100, Integer?)
        Dim n2 As Object = CType(Nothing, Integer?)
        Console.WriteLine(TypeOf n1 Is Integer & "|" & (n2 Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_typeof_is_array_type() {
    let src = r#"
Module Program
    Sub Main()
        Dim intArr As Object = New Integer() {1, 2, 3}
        Dim strArr As Object = New String() {"A", "B"}
        Console.WriteLine(TypeOf intArr Is Integer() & "|" & TypeOf strArr Is Array)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_typeof_is_delegate_type() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim act As Object = CType(Sub() Console.WriteLine("Hi"), Action)
        Console.WriteLine(TypeOf act Is Action & "|" & TypeOf act Is Delegate)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_typeof_is_enum_underlying_type() {
    let src = r#"
Imports System

Enum Level As Byte
    Low = 1
    High = 2
End Enum

Module Program
    Sub Main()
        Dim e As Object = Level.Low
        Console.WriteLine(TypeOf e Is Level & "|" & TypeOf e Is Enum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_typeof_is_struct_type() {
    let src = r#"
Imports System

Structure Point3D
    Public X, Y, Z As Double
End Structure

Module Program
    Sub Main()
        Dim pt As Object = New Point3D With {.X = 1}
        Console.WriteLine(TypeOf pt Is Point3D & "|" & TypeOf pt Is ValueType)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_is_operator_boxed_value_types_have_different_references() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = 100
        Dim box1 As Object = val
        Dim box2 As Object = val
        Console.WriteLine(box1 Is box2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_typeof_is_expression_pattern_in_if_statement() {
    let src = r#"
Module Program
    Sub Main()
        Dim item As Object = "Hello"
        If TypeOf item Is String Then
            Dim s As String = DirectCast(item, String)
            Console.WriteLine(s.Length)
        End If
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}
