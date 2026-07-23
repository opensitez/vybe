use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Anonymous Types Generic Inference & Expression Trees
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_anonymous_type_passed_to_generic_method() {
    let src = r#"
Module Program
    Private Function GetPropSummary(Of T)(item As T) As String
        Return item.GetType().Name
    End Function

    Sub Main()
        Dim obj = New With {.Name = "Test", .Value = 100}
        Console.WriteLine(GetPropSummary(obj).Contains("AnonymousType"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_anonymous_type_generic_list_creation() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Function CreateList(Of T)(ParamArray items As T()) As List(Of T)
        Return New List(Of T)(items)
    End Function

    Sub Main()
        Dim item1 = New With {.ID = 1}
        Dim item2 = New With {.ID = 2}
        Dim list = CreateList(item1, item2)
        Console.WriteLine(list.Count & ":" & list(0).ID & "," & list(1).ID)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2:1,2"]);
}

#[test]
fn test_vb_anonymous_type_type_inference_matching_properties() {
    let src = r#"
Module Program
    Private Sub ProcessPair(Of T)(first As T, second As T)
        Console.WriteLine("Types match successfully")
    End Sub

    Sub Main()
        Dim o1 = New With {.Code = "A", .Count = 10}
        Dim o2 = New With {.Code = "B", .Count = 20}
        ProcessPair(o1, o2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Types match successfully"]);
}

#[test]
fn test_vb_anonymous_type_key_order_determines_type_identity() {
    let src = r#"
Module Program
    Sub Main()
        Dim o1 = New With {.A = 1, .B = "X"}
        Dim o2 = New With {.A = 2, .B = "Y"}
        Console.WriteLine(o1.GetType() Is o2.GetType())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_anonymous_type_different_property_order_different_types() {
    let src = r#"
Module Program
    Sub Main()
        Dim o1 = New With {.A = 1, .B = "X"}
        Dim o2 = New With {.B = "X", .A = 1}
        Console.WriteLine(o1.GetType() Is o2.GetType())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_anonymous_type_key_vs_nonkey_determines_type_identity() {
    let src = r#"
Module Program
    Sub Main()
        Dim o1 = New With {Key .A = 1, .B = "X"}
        Dim o2 = New With {.A = 1, .B = "X"}
        Console.WriteLine(o1.GetType() Is o2.GetType())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_anonymous_type_generic_dictionary_values() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Private Function CreateDict(Of TKey, TValue)(k As TKey, v As TValue) As Dictionary(Of TKey, TValue)
        Dim d As New Dictionary(Of TKey, TValue)()
        d(k) = v
        Return d
    End Function

    Sub Main()
        Dim valObj = New With {.Status = "OK"}
        Dim dict = CreateDict("Key1", valObj)
        Console.WriteLine(dict("Key1").Status)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OK"]);
}

#[test]
fn test_vb_anonymous_type_linq_aggregate_projection() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20, 30}
        Dim stats = numbers.Aggregate(
            New With {.Sum = 0, .Count = 0},
            Function(acc, n) New With {.Sum = acc.Sum + n, .Count = acc.Count + 1}
        )
        Console.WriteLine("Sum=" & stats.Sum & "|Count=" & stats.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Sum=60|Count=3"]);
}

#[test]
fn test_vb_anonymous_type_expression_tree_building() {
    let src = r#"
Imports System.Linq.Expressions

Module Program
    Sub Main()
        Dim param = Expression.Parameter(GetType(String), "s")
        Dim anonExpr = Expression.New(
            GetType(Object)
        )
        Console.WriteLine(param.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["s"]);
}

#[test]
fn test_vb_anonymous_type_reflection_property_value_retrieval() {
    let src = r#"
Module Program
    Private Function GetPropValue(obj As Object, propName As String) As Object
        Dim prop = obj.GetType().GetProperty(propName)
        Return prop.GetValue(obj)
    End Function

    Sub Main()
        Dim anon = New With {.Title = "ReflectedTitle"}
        Console.WriteLine(GetPropValue(anon, "Title"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ReflectedTitle"]);
}

#[test]
fn test_vb_anonymous_type_generic_extension_method_invocation() {
    let src = r#"
Imports System.Runtime.CompilerServices

Module GenericExtensions
    <Extension()>
    Public Function ToJsonLikeString(Of T)(obj As T) As String
        Return obj.ToString()
    End Function
End Module

Module Program
    Sub Main()
        Dim item = New With {.ID = 10, .Name = "Item10"}
        Console.WriteLine(item.ToJsonLikeString().Contains("ID = 10"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_anonymous_type_array_of_interfaces_projection() {
    let src = r#"
Imports System.Linq

Interface IIdentifiable
    ReadOnly Property ID As Integer
End Interface

Class Item
    Implements IIdentifiable
    Public ReadOnly Property ID As Integer Implements IIdentifiable.ID
    Public Sub New(id As Integer) : Me.ID = id : End Sub
End Class

Module Program
    Sub Main()
        Dim items As IIdentifiable() = {New Item(1), New Item(2)}
        Dim projected = items.Select(Function(i) New With {.ItemID = i.ID})
        For Each p In projected
            Console.WriteLine("ID:" & p.ItemID)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ID:1", "ID:2"]);
}

#[test]
fn test_vb_anonymous_type_property_type_inference_from_null() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj = New With {.Data = CType(Nothing, String)}
        Console.WriteLine(obj.Data Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_anonymous_type_property_type_inference_from_delegate() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim obj = New With {.Action = CType(Sub() Console.WriteLine("ActionInAnon"), Action)}
        obj.Action()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ActionInAnon"]);
}

#[test]
fn test_vb_anonymous_type_property_type_inference_from_tuple() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj = New With {.Pair = (10, "Ten")}
        Console.WriteLine(obj.Pair.Item1 & "=" & obj.Pair.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10=Ten"]);
}

#[test]
fn test_vb_anonymous_type_property_type_inference_from_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj = New With {.Numbers = {1, 2, 3}}
        Console.WriteLine(String.Join(",", obj.Numbers))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_anonymous_type_nested_in_generic_class() {
    let src = r#"
Class Wrapper(Of T)
    Public Function Wrap(val As T) As Object
        Return New With {.WrappedVal = val}
    End Function
End Class

Module Program
    Sub Main()
        Dim w As New Wrapper(Of Double)()
        Dim anon As Dynamic = w.Wrap(3.14159)
        Console.WriteLine(anon.WrappedVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3.14159"]);
}

#[test]
fn test_vb_anonymous_type_as_method_return_value_via_object() {
    let src = r#"
Module Program
    Private Function GetAnonData() As Object
        Return New With {.Status = "Ready", .Code = 200}
    End Function

    Sub Main()
        Dim data As Object = GetAnonData()
        Dim propStatus = data.GetType().GetProperty("Status")
        Console.WriteLine(propStatus.GetValue(data))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Ready"]);
}

#[test]
fn test_vb_anonymous_type_equals_null_comparison() {
    let src = r#"
Module Program
    Sub Main()
        Dim obj = New With {.Name = "Test"}
        Console.WriteLine(obj.Equals(Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_anonymous_type_multiple_generic_type_parameters_inferred() {
    let src = r#"
Module Program
    Private Function PairUp(Of T1, T2)(a As T1, b As T2) As Object
        Return New With {.First = a, .Second = b}
    End Function

    Sub Main()
        Dim pair As Object = PairUp(100, "Hundred")
        Dim fProp = pair.GetType().GetProperty("First")
        Dim sProp = pair.GetType().GetProperty("Second")
        Console.WriteLine(fProp.GetValue(pair) & "=" & sProp.GetValue(pair))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100=Hundred"]);
}
