use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: KeyValuePair(Of TKey, TValue) Struct Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_key_value_pair_construction_properties() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim kv As New KeyValuePair(Of String, Integer)("Age", 30)
        Console.WriteLine(kv.Key & "=" & kv.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Age=30"]);
}

#[test]
fn test_vb_key_value_pair_create_factory_method() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim kv = KeyValuePair.Create("Status", True)
        Console.WriteLine(kv.Key & ":" & kv.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Status:True"]);
}

#[test]
fn test_vb_key_value_pair_deconstruct() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim kv As New KeyValuePair(Of String, Double)("Price", 19.99)
        Dim key As String = Nothing
        Dim val As Double = 0
        kv.Deconstruct(key, val)
        Console.WriteLine(key & " --> " & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Price --> 19.99"]);
}

#[test]
fn test_vb_key_value_pair_to_string_representation() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim kv As New KeyValuePair(Of Integer, String)(1, "One")
        Console.WriteLine(kv.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[1, One]"]);
}

#[test]
fn test_vb_key_value_pair_array_iteration() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim pairs As KeyValuePair(Of String, Integer)() = {
            New KeyValuePair(Of String, Integer)("A", 1),
            New KeyValuePair(Of String, Integer)("B", 2)
        }
        For Each pair In pairs
            Console.WriteLine(pair.Key & ":" & pair.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A:1", "B:2"]);
}

#[test]
fn test_vb_key_value_pair_list_filtering() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim list As New List(Of KeyValuePair(Of String, Integer)) From {
            New KeyValuePair(Of String, Integer)("High", 90),
            New KeyValuePair(Of String, Integer)("Low", 10),
            New KeyValuePair(Of String, Integer)("High", 85)
        }
        Dim highs = list.Where(Function(kv) kv.Key = "High")
        For Each kv In highs
            Console.WriteLine(kv.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["90", "85"]);
}

#[test]
fn test_vb_key_value_pair_value_type_semantics() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim kv1 As New KeyValuePair(Of String, Integer)("K", 10)
        Dim kv2 As KeyValuePair(Of String, Integer) = kv1
        Console.WriteLine(kv1.Key = kv2.Key AndAlso kv1.Value = kv2.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_key_value_pair_null_key_and_null_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim kv As New KeyValuePair(Of String, String)(Nothing, Nothing)
        Console.WriteLine(kv.Key Is Nothing & "|" & kv.Value Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_key_value_pair_complex_object_value() {
    let src = r#"
Imports System.Collections.Generic

Class Person
    Public Property Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
End Class

Module Program
    Sub Main()
        Dim kv As New KeyValuePair(Of Integer, Person)(101, New Person("Alice"))
        Console.WriteLine(kv.Key & "=>" & kv.Value.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["101=>Alice"]);
}

#[test]
fn test_vb_key_value_pair_struct_key_and_value() {
    let src = r#"
Imports System.Collections.Generic

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim kv As New KeyValuePair(Of Point, Point)(New Point(0, 0), New Point(10, 20))
        Console.WriteLine(kv.Key.X & " to " & kv.Value.X & "," & kv.Value.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0 to 10,20"]);
}

#[test]
fn test_vb_key_value_pair_dictionary_to_array_conversion() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"X", 10}, {"Y", 20}}
        Dim arr As KeyValuePair(Of String, Integer)() = dict.ToArray()
        Console.WriteLine(arr(0).Key & "=" & arr(0).Value & "|" & arr(1).Key & "=" & arr(1).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["X=10|Y=20"]);
}

#[test]
fn test_vb_key_value_pair_tuple_conversions() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim kv As New KeyValuePair(Of String, Integer)("Num", 7)
        Dim tuple As (String, Integer) = (kv.Key, kv.Value)
        Console.WriteLine(tuple.Item1 & ":" & tuple.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Num:7"]);
}

#[test]
fn test_vb_key_value_pair_enum_types() {
    let src = r#"
Imports System.Collections.Generic

Enum Priority
    Low
    High
End Enum

Module Program
    Sub Main()
        Dim kv As New KeyValuePair(Of Priority, Priority)(Priority.Low, Priority.High)
        Console.WriteLine(kv.Key.ToString() & "->" & kv.Value.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Low->High"]);
}

#[test]
fn test_vb_key_value_pair_in_custom_collection() {
    let src = r#"
Imports System.Collections.Generic

Class Cache(Of K, V)
    Private items As New List(Of KeyValuePair(Of K, V))()
    Public Sub Put(k As K, v As V)
        items.Add(New KeyValuePair(Of K, V)(k, v))
    End Sub
    Public Function GetFirst() As KeyValuePair(Of K, V)
        Return items(0)
    End Function
End Class

Module Program
    Sub Main()
        Dim c As New Cache(Of String, String)()
        c.Put("Token", "ABC123")
        Dim kv = c.GetFirst()
        Console.WriteLine(kv.Key & "=" & kv.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Token=ABC123"]);
}

#[test]
fn test_vb_key_value_pair_nested_pair() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim inner As New KeyValuePair(Of String, Integer)("Inner", 42)
        Dim outer As New KeyValuePair(Of String, KeyValuePair(Of String, Integer))("Outer", inner)
        Console.WriteLine(outer.Key & "->" & outer.Value.Key & "->" & outer.Value.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Outer->Inner->42"]);
}

#[test]
fn test_vb_key_value_pair_equality_comparison() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim p1 As New KeyValuePair(Of String, Integer)("A", 1)
        Dim p2 As New KeyValuePair(Of String, Integer)("A", 1)
        Dim p3 As New KeyValuePair(Of String, Integer)("A", 2)
        Console.WriteLine(p1.Equals(p2) & "|" & p1.Equals(p3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_key_value_pair_hash_code() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim p1 As New KeyValuePair(Of String, Integer)("A", 1)
        Dim p2 As New KeyValuePair(Of String, Integer)("A", 1)
        Console.WriteLine(p1.GetHashCode() = p2.GetHashCode())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_key_value_pair_datetime_key() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim kv As New KeyValuePair(Of DateTime, String)(New DateTime(2025, 5, 1), "Labor Day")
        Console.WriteLine(kv.Key.ToString("yyyy-MM-dd") & ":" & kv.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-05-01:Labor Day"]);
}

#[test]
fn test_vb_key_value_pair_guid_key() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim g = Guid.Parse("11111111-2222-3333-4444-555555555555")
        Dim kv As New KeyValuePair(Of Guid, String)(g, "SessionData")
        Console.WriteLine(kv.Key.ToString() & "|" & kv.Value)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["11111111-2222-3333-4444-555555555555|SessionData"]
    );
}

#[test]
fn test_vb_key_value_pair_collection_projection() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim keys = {"K1", "K2", "K3"}
        Dim vals = {10, 20, 30}
        Dim pairs = keys.Zip(vals, Function(k, v) New KeyValuePair(Of String, Integer)(k, v))
        For Each p In pairs
            Console.WriteLine(p.Key & "=" & p.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["K1=10", "K2=20", "K3=30"]);
}
