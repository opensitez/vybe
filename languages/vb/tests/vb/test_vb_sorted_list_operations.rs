use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: SortedList(Of K, V) Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_sorted_list_auto_sorting() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New SortedList(Of String, Integer)
        list.Add("Zebra", 26)
        list.Add("Apple", 1)
        list.Add("Monkey", 13)
        For Each kvp In list
            Console.WriteLine(kvp.Key & ":" & kvp.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Apple:1", "Monkey:13", "Zebra:26"]);
}

#[test]
fn test_vb_sorted_list_index_access_by_key_and_position() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New SortedList(Of Integer, String) From {{30, "Thirty"}, {10, "Ten"}, {20, "Twenty"}}
        Console.WriteLine(list(20))
        Console.WriteLine(list.Keys(0))
        Console.WriteLine(list.Values(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Twenty", "10", "Ten"]);
}

#[test]
fn test_vb_sorted_list_index_of_key_and_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New SortedList(Of String, Integer) From {{"A", 100}, {"B", 200}, {"C", 300}}
        Dim keyIdx As Integer = list.IndexOfKey("B")
        Dim valIdx As Integer = list.IndexOfValue(300)
        Console.WriteLine(keyIdx)
        Console.WriteLine(valIdx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1", "2"]);
}

#[test]
fn test_vb_sorted_list_remove_at_index() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New SortedList(Of Integer, String) From {{1, "One"}, {2, "Two"}, {3, "Three"}}
        list.RemoveAt(1)
        Console.WriteLine(String.Join(",", list.Values))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["One,Three"]);
}

#[test]
fn test_vb_sorted_list_capacity_trim_excess() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New SortedList(Of Integer, Integer)(100)
        list.Add(1, 10)
        list.TrimExcess()
        Console.WriteLine(list.Capacity >= 1)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_sorted_list_set_value_at_index() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New SortedList(Of Integer, String) From {{10, "X"}, {20, "Y"}}
        list.SetValueAt(1, "Z")
        Console.WriteLine(list(20))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Z"]);
}

#[test]
fn test_vb_sorted_list_contains_key_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New SortedList(Of String, String) From {{"K1", "V1"}}
        Console.WriteLine(list.ContainsKey("K1"))
        Console.WriteLine(list.ContainsValue("V1"))
        Console.WriteLine(list.ContainsKey("K2"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True", "False"]);
}

#[test]
fn test_vb_sorted_list_try_get_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New SortedList(Of Integer, String) From {{1, "A"}}
        Dim outVal As String = Nothing
        Dim ok As Boolean = list.TryGetValue(1, outVal)
        Console.WriteLine(ok)
        Console.WriteLine(outVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "A"]);
}
