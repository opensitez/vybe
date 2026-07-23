use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array Binary Search Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_binary_search_found() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30, 40, 50}
        Dim idx As Integer = Array.BinarySearch(arr, 30)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_array_binary_search_not_found_bitwise_complement() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30, 40, 50}
        Dim idx As Integer = Array.BinarySearch(arr, 25)
        Console.WriteLine(idx < 0)
        Console.WriteLine(Not idx) ' Index where 25 would be inserted (2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "2"]);
}

#[test]
fn test_vb_array_binary_search_range() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {5, 10, 15, 20, 25, 30}
        Dim idx As Integer = Array.BinarySearch(arr, 1, 4, 20)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_array_binary_search_custom_comparer() {
    let src = r#"
Imports System.Collections.Generic

Class DescendingComparer
    Implements IComparer(Of Integer)
    Public Function Compare(x As Integer, y As Integer) As Integer Implements IComparer(Of Integer).Compare
        Return y.CompareTo(x)
    End Function
End Class

Module Program
    Sub Main()
        Dim arr As Integer() = {50, 40, 30, 20, 10}
        Dim idx As Integer = Array.BinarySearch(arr, 40, New DescendingComparer())
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_array_binary_search_string_ordinal() {
    let src = r#"
Module Program
    Sub Main()
        Dim words As String() = {"apple", "banana", "cherry", "date"}
        Dim idx As Integer = Array.BinarySearch(words, "cherry")
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_array_binary_search_first_element() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {100, 200, 300}
        Dim idx As Integer = Array.BinarySearch(arr, 100)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_array_binary_search_last_element() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {100, 200, 300}
        Dim idx As Integer = Array.BinarySearch(arr, 300)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_array_binary_search_before_first() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        Dim idx As Integer = Array.BinarySearch(arr, 5)
        Console.WriteLine(Not idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_array_binary_search_after_last() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        Dim idx As Integer = Array.BinarySearch(arr, 35)
        Console.WriteLine(Not idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_array_binary_search_range_custom_comparer() {
    let src = r#"
Module Program
    Sub Main()
        Dim words As String() = {"a", "B", "c", "D"}
        Array.Sort(words, StringComparer.OrdinalIgnoreCase)
        Dim idx As Integer = Array.BinarySearch(words, 0, 4, "b", StringComparer.OrdinalIgnoreCase)
        Console.WriteLine(idx >= 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
