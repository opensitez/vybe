use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: SortedDictionary(Of K, V) Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_sorted_dict_auto_sorting_keys() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)
        dict.Add(30, "Thirty")
        dict.Add(10, "Ten")
        dict.Add(20, "Twenty")
        For Each kvp In dict
            Console.WriteLine(kvp.Key & ":" & kvp.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10:Ten", "20:Twenty", "30:Thirty"]);
}

#[test]
fn test_vb_sorted_dict_custom_comparer() {
    let src = r#"
Imports System.Collections.Generic

Class DescendingIntComparer
    Implements IComparer(Of Integer)
    Public Function Compare(x As Integer, y As Integer) As Integer Implements IComparer(Of Integer).Compare
        Return y.CompareTo(x)
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)(New DescendingIntComparer())
        dict(1) = "One"
        dict(3) = "Three"
        dict(2) = "Two"
        For Each kvp In dict
            Console.WriteLine(kvp.Key)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "2", "1"]);
}

#[test]
fn test_vb_sorted_dict_contains_key_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer) From {{"B", 2}, {"A", 1}}
        Console.WriteLine(dict.ContainsKey("A"))
        Console.WriteLine(dict.ContainsValue(2))
        Console.WriteLine(dict.ContainsKey("C"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True", "False"]);
}

#[test]
fn test_vb_sorted_dict_try_get_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, String) From {{"x", "100"}}
        Dim val As String = Nothing
        Dim found As Boolean = dict.TryGetValue("x", val)
        Console.WriteLine(found)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "100"]);
}

#[test]
fn test_vb_sorted_dict_remove_key() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String) From {{1, "A"}, {2, "B"}}
        Dim ok As Boolean = dict.Remove(1)
        Console.WriteLine(ok)
        Console.WriteLine(dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "1"]);
}

#[test]
fn test_vb_sorted_dict_keys_values_collections() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String) From {{3, "C"}, {1, "A"}, {2, "B"}}
        Console.WriteLine(String.Join(",", dict.Keys))
        Console.WriteLine(String.Join(",", dict.Values))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3", "A,B,C"]);
}

#[test]
fn test_vb_sorted_dict_clear() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, Integer) From {{1, 10}, {2, 20}}
        dict.Clear()
        Console.WriteLine(dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_sorted_dict_indexer_get_set() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)
        dict("Key") = 42
        dict("Key") += 8
        Console.WriteLine(dict("Key"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}
