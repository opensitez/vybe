use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: IIf Legacy Function vs If Ternary Operator Semantics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_iif_eager_evaluation_both_branches() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Private Function SideEffect(msg As String, val As Integer) As Integer
        Console.WriteLine("Effect:" & msg)
        Return val
    End Function

    Sub Main()
        ' IIf eagerly evaluates both truepart and falsepart!
        Dim res = IIf(True, SideEffect("TrueBranch", 10), SideEffect("FalseBranch", 20))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Effect:TrueBranch", "Effect:FalseBranch", "10"]
    );
}

#[test]
fn test_vb_if_operator_short_circuit_evaluation() {
    let src = r#"
Module Program
    Private Function SideEffect(msg As String, val As Integer) As Integer
        Console.WriteLine("Effect:" & msg)
        Return val
    End Function

    Sub Main()
        ' If ternary operator short-circuits (only true branch evaluated)!
        Dim res = If(True, SideEffect("TrueBranch", 10), SideEffect("FalseBranch", 20))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Effect:TrueBranch", "10"]);
}

#[test]
fn test_vb_iif_returns_object_requiring_conversion() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        ' IIf returns Object type
        Dim objRes = IIf(1 > 0, 100, 200)
        Dim num As Integer = CInt(objRes)
        Console.WriteLine(num)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_if_operator_type_inference() {
    let src = r#"
Module Program
    Sub Main()
        ' If operator infers exact common type (Integer)
        Dim num = If(1 > 0, 100, 200)
        Console.WriteLine(num + 50)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["150"]);
}

#[test]
fn test_vb_if_binary_coalesce_operator() {
    let src = r#"
Module Program
    Sub Main()
        Dim str1 As String = Nothing
        Dim str2 As String = "Fallback"
        Dim res = If(str1, str2)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Fallback"]);
}

#[test]
fn test_vb_if_binary_coalesce_first_not_null() {
    let src = r#"
Module Program
    Sub Main()
        Dim str1 As String = "Primary"
        Dim str2 As String = "Fallback"
        Dim res = If(str1, str2)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Primary"]);
}

#[test]
fn test_vb_iif_divide_by_zero_side_effect_throws() {
    let src = r#"
Imports System
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Try
            ' Because IIf evaluates both branches, 10 / 0 in falsepart throws DivideByZeroException even when condition is True!
            Dim res = IIf(True, 42, 10 \ 0)
        Catch ex As DivideByZeroException
            Console.WriteLine("DivideByZeroException Caught in Eager IIf")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["DivideByZeroException Caught in Eager IIf"]
    );
}

#[test]
fn test_vb_if_operator_divide_by_zero_guarded_safe() {
    let src = r#"
Module Program
    Sub Main()
        Dim divisor = 0
        ' If ternary short-circuits so false branch 10 \ divisor is NOT evaluated!
        Dim res = If(divisor <> 0, 10 \ divisor, -1)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1"]);
}

#[test]
fn test_vb_if_binary_coalesce_nullable_integer() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim n1 As Integer? = Nothing
        Dim n2 As Integer? = 77
        Dim res = If(n1, n2.Value)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["77"]);
}

#[test]
fn test_vb_nested_if_ternary_operators() {
    let src = r#"
Module Program
    Sub Main()
        Dim score = 85
        Dim grade = If(score >= 90, "A", If(score >= 80, "B", "C"))
        Console.WriteLine(grade)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B"]);
}

#[test]
fn test_vb_if_operator_string_concatenation() {
    let src = r#"
Module Program
    Sub Main()
        Dim name As String = "Bob"
        Dim msg = "Hello " & If(name IsNot Nothing, name, "Guest")
        Console.WriteLine(msg)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello Bob"]);
}

#[test]
fn test_vb_if_operator_value_type_conversion() {
    let src = r#"
Module Program
    Sub Main()
        ' Common type of Integer and Double is Double
        Dim res = If(True, 5, 2.5)
        Console.WriteLine(res.GetType().Name & ":" & res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Double:5"]);
}

#[test]
fn test_vb_iif_with_null_arguments() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Dim res = IIf(False, "NotNull", Nothing)
        Console.WriteLine(res Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_if_operator_boolean_condition_expression() {
    let src = r#"
Module Program
    Sub Main()
        Dim a = 10
        Dim b = 20
        Dim maxVal = If(a > b, a, b)
        Console.WriteLine(maxVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_if_operator_returning_lambda() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim flag = True
        Dim op As Func(Of Integer, Integer) = If(flag, Function(x) x * 2, Function(x) x * 3)
        Console.WriteLine(op(10))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_if_binary_coalesce_chained() {
    let src = r#"
Module Program
    Sub Main()
        Dim v1 As String = Nothing
        Dim v2 As String = Nothing
        Dim v3 As String = "Final"
        Dim res = If(v1, If(v2, v3))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Final"]);
}

#[test]
fn test_vb_iif_legacy_boolean_coercion_condition() {
    let src = r#"
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        ' IIf accepts numeric condition: non-zero is True, 0 is False
        Dim res1 = IIf(1, "TruePart", "FalsePart")
        Dim res2 = IIf(0, "TruePart", "FalsePart")
        Console.WriteLine(res1 & "|" & res2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TruePart|FalsePart"]);
}

#[test]
fn test_vb_if_operator_custom_class_inheritance_common_type() {
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
        Dim isDog = True
        ' Common base class Animal is inferred!
        Dim pet As Animal = If(isDog, DirectCast(New Dog(), Animal), DirectCast(New Cat(), Animal))
        Console.WriteLine(pet.GetType().Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Dog"]);
}

#[test]
fn test_vb_if_operator_structure_result() {
    let src = r#"
Structure Point2D
    Public X, Y As Integer
End Structure

Module Program
    Sub Main()
        Dim p1 As New Point2D With {.X = 1, .Y = 1}
        Dim p2 As New Point2D With {.X = 2, .Y = 2}
        Dim selected = If(True, p1, p2)
        Console.WriteLine(selected.X & "," & selected.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,1"]);
}

#[test]
fn test_vb_if_operator_null_reference_method_guard() {
    let src = r#"
Class Customer
    Public Property Name As String = "Alice"
End Class

Module Program
    Sub Main()
        Dim c As Customer = Nothing
        Dim name = If(c IsNot Nothing, c.Name, "Unknown")
        Console.WriteLine(name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Unknown"]);
}
