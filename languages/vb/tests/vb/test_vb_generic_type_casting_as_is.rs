use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generic Type Casting, TryCast, DirectCast & TypeOf
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_typeof_is_operator_check() {
    let src = r#"
Module Program
    Private Function IsType(Of T)(obj As Object) As Boolean
        Return TypeOf obj Is T
    End Function

    Sub Main()
        Console.WriteLine(IsType(Of Integer)(42) & "|" & IsType(Of String)(42))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_generic_trycast_reference_type_success() {
    let src = r#"
Module Program
    Private Function SafeCast(Of T As Class)(obj As Object) As T
        Return TryCast(obj, T)
    End Function

    Sub Main()
        Dim str = SafeCast(Of String)("Hello World")
        Console.WriteLine(str IsNot Nothing & "|" & str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|Hello World"]);
}

#[test]
fn test_vb_generic_trycast_reference_type_failure_returns_nothing() {
    let src = r#"
Module Program
    Private Function SafeCast(Of T As Class)(obj As Object) As T
        Return TryCast(obj, T)
    End Function

    Sub Main()
        Dim str = SafeCast(Of String)(12345)
        Console.WriteLine(str Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_generic_directcast_reference_type_success() {
    let src = r#"
Module Program
    Private Function CastDirect(Of T)(obj As Object) As T
        Return DirectCast(obj, T)
    End Function

    Sub Main()
        Dim str = CastDirect(Of String)("DirectCastValue")
        Console.WriteLine(str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["DirectCastValue"]);
}

#[test]
fn test_vb_generic_directcast_value_type_boxed_unboxing() {
    let src = r#"
Module Program
    Sub Main()
        Dim boxed As Object = 99
        Dim unboxed As Integer = DirectCast(boxed, Integer)
        Console.WriteLine(unboxed)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99"]);
}

#[test]
fn test_vb_generic_directcast_invalid_cast_exception() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim boxed As Object = "NotAnInt"
            Dim num As Integer = DirectCast(boxed, Integer)
            Console.WriteLine(num)
        Catch ex As InvalidCastException
            Console.WriteLine("DirectCast InvalidCastException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["DirectCast InvalidCastException Caught"]);
}

#[test]
fn test_vb_generic_ctype_conversion_numeric_promotion() {
    let src = r#"
Module Program
    Private Function ConvertGeneric(Of T)(obj As Object) As T
        Return CType(obj, T)
    End Function

    Sub Main()
        Dim doubleVal As Double = ConvertGeneric(Of Double)(42)
        Console.WriteLine(doubleVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_generic_typeof_isnot_operator_check() {
    let src = r#"
Module Program
    Private Function IsNotType(Of T)(obj As Object) As Boolean
        Return TypeOf obj IsNot T
    End Function

    Sub Main()
        Console.WriteLine(IsNotType(Of String)(100) & "|" & IsNotType(Of Integer)(100))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_generic_class_inheritance_upcasting() {
    let src = r#"
Class Animal : End Class
Class Dog : Inherits Animal : End Class

Module Program
    Private Function Upcast(Of TDerived As TBase, TBase As Class)(item As TDerived) As TBase
        Return DirectCast(CObj(item), TBase)
    End Function

    Sub Main()
        Dim d As New Dog()
        Dim a As Animal = Upcast(Of Dog, Animal)(d)
        Console.WriteLine(a IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_generic_interface_casting() {
    let src = r#"
Interface IService : End Interface
Class ServiceImpl : Implements IService : End Class

Module Program
    Private Function AsInterface(Of TInterface As Class)(obj As Object) As TInterface
        Return TryCast(obj, TInterface)
    End Function

    Sub Main()
        Dim impl As Object = New ServiceImpl()
        Dim svc As IService = AsInterface(Of IService)(impl)
        Console.WriteLine(svc IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_generic_enum_casting_from_int() {
    let src = r#"
Enum Level
    Low = 1
    High = 2
End Enum

Module Program
    Private Function EnumCast(Of TEnum As Structure)(val As Integer) As TEnum
        Return CType(CObj(val), TEnum)
    End Function

    Sub Main()
        Dim l As Level = EnumCast(Of Level)(2)
        Console.WriteLine(l.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["High"]);
}

#[test]
fn test_vb_generic_nullable_casting() {
    let src = r#"
Imports System

Module Program
    Private Function AsNullable(Of T As Structure)(obj As Object) As Nullable(Of T)
        If TypeOf obj Is T Then
            Return CType(obj, T)
        End If
        Return Nothing
    End Function

    Sub Main()
        Dim n1 = AsNullable(Of Integer)(50)
        Dim n2 = AsNullable(Of Integer)("NotInt")
        Console.WriteLine(n1.HasValue & ":" & n1.GetValueOrDefault() & "|" & n2.HasValue)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:50|False"]);
}

#[test]
fn test_vb_generic_typeof_is_expression_pattern() {
    let src = r#"
Module Program
    Private Function Inspect(Of T)(val As T) As String
        If TypeOf CObj(val) Is Integer Then Return "Integer"
        If TypeOf CObj(val) Is String Then Return "String"
        Return "Unknown"
    End Function

    Sub Main()
        Console.WriteLine(Inspect(10) & "|" & Inspect("ABC") & "|" & Inspect(3.14))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Integer|String|Unknown"]);
}

#[test]
fn test_vb_generic_type_of_exact_match_vs_derived() {
    let src = r#"
Class Parent : End Class
Class Child : Inherits Parent : End Class

Module Program
    Sub Main()
        Dim c As Object = New Child()
        Console.WriteLine((TypeOf c Is Parent) & "|" & (c.GetType() Is GetType(Parent)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_generic_struct_casting_to_object_boxing() {
    let src = r#"
Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Module Program
    Private Function BoxValue(Of T)(val As T) As Object
        Return CObj(val)
    End Function

    Sub Main()
        Dim p As New Point With {.X = 1, .Y = 2}
        Dim boxed = BoxValue(p)
        Console.WriteLine(TypeOf boxed Is Point)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_generic_array_type_casting() {
    let src = r#"
Module Program
    Private Function CastArray(Of T)(arr As Array) As T()
        Return CType(arr, T())
    End Function

    Sub Main()
        Dim rawArr As Array = New String() {"A", "B"}
        Dim strArr As String() = CastArray(Of String)(rawArr)
        Console.WriteLine(String.Join(",", strArr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A,B"]);
}

#[test]
fn test_vb_generic_delegate_casting() {
    let src = r#"
Imports System

Module Program
    Private Function CastDelegate(Of TDelegate As Class)(d As [Delegate]) As TDelegate
        Return TryCast(d, TDelegate)
    End Function

    Sub Main()
        Dim act As Action = Sub() Console.WriteLine("Action Executed")
        Dim castAct = CastDelegate(Of Action)(act)
        castAct()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Action Executed"]);
}

#[test]
fn test_vb_generic_is_operator_with_value_type_boxing_check() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = 100
        Dim obj As Object = val
        Console.WriteLine(TypeOf obj Is Integer & "|" & TypeOf obj Is Double)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_generic_tuple_type_casting() {
    let src = r#"
Module Program
    Sub Main()
        Dim tupleObj As Object = ("Key", 42)
        Console.WriteLine(TypeOf tupleObj Is (String, Integer))
        Dim tuple = CType(tupleObj, (String, Integer))
        Console.WriteLine(tuple.Item1 & "=" & tuple.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "Key=42"]);
}

#[test]
fn test_vb_generic_type_argument_reflection_equality() {
    let src = r#"
Module Program
    Private Function CheckTypeEquivalence(Of T)(obj As Object) As Boolean
        Return obj.GetType() Is GetType(T)
    End Function

    Sub Main()
        Console.WriteLine(CheckTypeEquivalence(Of Integer)(10) & "|" & CheckTypeEquivalence(Of Object)(10))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}
