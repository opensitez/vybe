use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: IsNot Operator & Compound Null Reference Guard Semantics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_isnot_operator_non_null_object() {
    let src = r#"
Module Program
    Class Node
    End Class

    Sub Main()
        Dim n As New Node()
        Console.WriteLine(n IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_isnot_operator_null_object() {
    let src = r#"
Module Program
    Class Node
    End Class

    Sub Main()
        Dim n As Node = Nothing
        Console.WriteLine(n IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_isnot_operator_reference_comparison() {
    let src = r#"
Module Program
    Class Item
    End Class

    Sub Main()
        Dim item1 As New Item()
        Dim item2 As New Item()
        Console.WriteLine(item1 IsNot item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_isnot_operator_same_reference_comparison() {
    let src = r#"
Module Program
    Class Item
    End Class

    Sub Main()
        Dim item1 As New Item()
        Dim item2 As Item = item1
        Console.WriteLine(item1 IsNot item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_isnot_operator_short_circuit_andalso() {
    let src = r#"
Module Program
    Class Container
        Public Property Content As String = "Data"
    End Class

    Sub Main()
        Dim c As Container = Nothing
        Dim hasData = (c IsNot Nothing) AndAlso (c.Content = "Data")
        Console.WriteLine(hasData)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_isnot_operator_typeof_combination() {
    let src = r#"
Module Program
    Class BaseType
    End Class

    Class SubType
        Inherits BaseType
    End Class

    Sub Main()
        Dim obj As BaseType = New BaseType()
        Console.WriteLine(TypeOf obj IsNot SubType)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_isnot_operator_string_empty_vs_nothing() {
    let src = r#"
Module Program
    Sub Main()
        Dim emptyStr As String = ""
        Dim nullStr As String = Nothing
        Console.WriteLine((emptyStr IsNot Nothing) & "|" & (nullStr IsNot Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_isnot_operator_boxed_value_types() {
    let src = r#"
Module Program
    Sub Main()
        Dim b1 As Object = 100
        Dim b2 As Object = 100
        ' Boxed value types have distinct reference identities!
        Console.WriteLine(b1 IsNot b2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_isnot_operator_nullable_type_has_value_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim n1 As Integer? = 42
        Dim n2 As Integer? = Nothing
        ' In VB.NET, Nullable types compared with IsNot Nothing check HasValue!
        Console.WriteLine((n1 IsNot Nothing) & "|" & (n2 IsNot Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_isnot_operator_delegate_subscription_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim act As Action = Sub() Console.WriteLine("Action")
        Dim nullAct As Action = Nothing
        Console.WriteLine((act IsNot Nothing) & "|" & (nullAct IsNot Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_isnot_operator_array_instantiation_check() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr1 As Integer() = New Integer() {1, 2}
        Dim arr2 As Integer() = Nothing
        Console.WriteLine((arr1 IsNot Nothing) & "|" & (arr2 IsNot Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_isnot_operator_in_while_loop_guard() {
    let src = r#"
Module Program
    Class Node
        Public Value As Integer
        Public NextNode As Node
    End Class

    Sub Main()
        Dim head As New Node With {.Value = 1, .NextNode = New Node With {.Value = 2}}
        Dim current As Node = head
        Dim sum = 0
        While current IsNot Nothing
            sum += current.Value
            current = current.NextNode
        End While
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_isnot_operator_in_if_else_chain() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj As Object = "Sample"
        If obj IsNot Nothing Then
            Console.WriteLine("Object Exists: " & obj.ToString())
        Else
            Console.WriteLine("Object Is Null")
        End If
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Object Exists: Sample"]);
}

#[test]
fn test_vb_isnot_operator_with_cint_coercion() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Object = 50
        Console.WriteLine(val IsNot Nothing AndAlso CInt(val) > 25)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_isnot_operator_generic_class_type_check() {
    let src = r#"
Class Wrapper(Of T)
    Public Item As T
End Class

Module Program
    Sub Main()
        Dim w As New Wrapper(Of String) With {.Item = "Data"}
        Console.WriteLine(w IsNot Nothing AndAlso w.Item IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_isnot_operator_expression_in_ternary_if() {
    let src = r#"
Module Program
    Sub Main()
        Dim str As String = "Active"
        Dim status = If(str IsNot Nothing, str.ToUpper(), "OFFLINE")
        Console.WriteLine(status)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ACTIVE"]);
}

#[test]
fn test_vb_isnot_operator_event_handler_null_guard() {
    let src = r#"
Imports System

Module Program
    Public Event CustomEvent As EventHandler

    Sub Main()
        ' In VB.NET CustomEventEvent field can be checked for IsNot Nothing before raising
        Console.WriteLine(CustomEventEvent IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_isnot_operator_struct_value_type_cannot_be_null() {
    let src = r#"
Structure SimplePoint
    Public X As Integer
End Structure

Module Program
    Sub Main()
        Dim p As New SimplePoint()
        ' Value types boxed to object can be checked against Nothing
        Dim objP As Object = p
        Console.WriteLine(objP IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_isnot_operator_multiple_null_checks() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Object = "A"
        Dim b As Object = "B"
        Dim c As Object = Nothing
        Console.WriteLine(a IsNot Nothing AndAlso b IsNot Nothing AndAlso c IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_isnot_operator_nested_class_properties() {
    let src = r#"
Module Program
    Class Engine
        Public Property Core As CoreUnit
    End Class

    Class CoreUnit
        Public Property Code As String = "C100"
    End Class

    Sub Main()
        Dim eng As New Engine()
        Console.WriteLine(eng IsNot Nothing AndAlso eng.Core IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}
