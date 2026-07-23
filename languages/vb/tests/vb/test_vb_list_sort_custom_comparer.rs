use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: List(Of T) Sort & Search with Custom Comparers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_list_sort_default() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {5, 2, 8, 1, 9}
        list.Sort()
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,5,8,9"]);
}

#[test]
fn test_vb_list_sort_comparison_lambda() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"apple", "hi", "banana", "cat"}
        list.Sort(Function(a, b) a.Length.CompareTo(b.Length))
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["hi,cat,apple,banana"]);
}

#[test]
fn test_vb_list_sort_range() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {5, 4, 3, 2, 1}
        list.Sort(1, 3, Nothing)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5,2,3,4,1"]);
}

#[test]
fn test_vb_list_binary_search_found() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {10, 20, 30, 40, 50}
        Dim idx As Integer = list.BinarySearch(30)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_list_convert_all_projection() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 3}
        Dim strings As List(Of String) = list.ConvertAll(Function(x) "Num: " & x)
        Console.WriteLine(String.Join(",", strings))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Num: 1,Num: 2,Num: 3"]);
}

#[test]
fn test_vb_list_true_for_all_predicate() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {2, 4, 6, 8}
        Console.WriteLine(list.TrueForAll(Function(x) x Mod 2 = 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_list_find_index_matching() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"one", "two", "three", "four"}
        Dim idx As Integer = list.FindIndex(Function(s) s.StartsWith("t"))
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_list_find_last_index_matching() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"one", "two", "three", "four"}
        Dim idx As Integer = list.FindLastIndex(Function(s) s.StartsWith("t"))
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_list_remove_all_predicate() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 3, 4, 5, 6}
        Dim count As Integer = list.RemoveAll(Function(x) x Mod 2 = 0)
        Console.WriteLine(count)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "1,3,5"]);
}

#[test]
fn test_vb_list_as_read_only_wrapper() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"A", "B", "C"}
        Dim ro As ReadOnlyCollection(Of String) = list.AsReadOnly()
        Console.WriteLine(ro.Count)
        Console.WriteLine(ro(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "B"]);
}
