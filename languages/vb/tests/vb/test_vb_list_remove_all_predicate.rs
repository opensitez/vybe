use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: List(Of T).RemoveAll Predicates & Collection Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_list_remove_all_evens() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 3, 4, 5, 6}
        Dim removedCount As Integer = list.RemoveAll(Function(n) n Mod 2 = 0)
        Console.WriteLine(removedCount)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "1,3,5"]);
}

#[test]
fn test_vb_list_remove_all_no_match() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 3, 5}
        Dim removedCount As Integer = list.RemoveAll(Function(n) n Mod 2 = 0)
        Console.WriteLine(removedCount)
        Console.WriteLine(list.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0", "3"]);
}

#[test]
fn test_vb_list_remove_all_all_matches() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {2, 4, 6}
        Dim removedCount As Integer = list.RemoveAll(Function(n) n Mod 2 = 0)
        Console.WriteLine(removedCount)
        Console.WriteLine(list.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "0"]);
}

#[test]
fn test_vb_list_remove_all_string_predicates() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim words As New List(Of String) From {"apple", "banana", "apricot", "cherry"}
        words.RemoveAll(Function(w) w.StartsWith("a"))
        Console.WriteLine(String.Join(",", words))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["banana,cherry"]);
}

#[test]
fn test_vb_list_remove_all_complex_objects() {
    let src = r#"
Imports System.Collections.Generic

Class TaskItem
    Public Property Title As String
    Public Property IsDone As Boolean
    Public Sub New(t As String, done As Boolean)
        Title = t : IsDone = done
    End Sub
End Class

Module Program
    Sub Main()
        Dim tasks As New List(Of TaskItem) From {
            New TaskItem("T1", True),
            New TaskItem("T2", False),
            New TaskItem("T3", True)
        }
        tasks.RemoveAll(Function(t) t.IsDone)
        Console.WriteLine(tasks.Count & ":" & tasks(0).Title)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:T2"]);
}

#[test]
fn test_vb_list_remove_at_first_middle_last() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"A", "B", "C", "D", "E"}
        list.RemoveAt(0) ' Removes A
        list.RemoveAt(1) ' Removes C
        list.RemoveAt(list.Count - 1) ' Removes E
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B,D"]);
}

#[test]
fn test_vb_list_remove_range() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {10, 20, 30, 40, 50}
        list.RemoveRange(1, 3) ' Remove 20, 30, 40
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,50"]);
}

#[test]
fn test_vb_list_remove_range_all_elements() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 3}
        list.RemoveRange(0, 3)
        Console.WriteLine(list.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_list_remove_value_returns_bool() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"X", "Y", "Z"}
        Dim ok1 As Boolean = list.Remove("Y")
        Dim ok2 As Boolean = list.Remove("Missing")
        Console.WriteLine(ok1 & "|" & ok2 & "|" & String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False|X,Z"]);
}

#[test]
fn test_vb_list_remove_first_occurrence_only() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 2, 3, 2}
        list.Remove(2)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,2"]);
}

#[test]
fn test_vb_list_clear_resets_count_preserves_capacity() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer)(50) From {1, 2, 3, 4, 5}
        list.Clear()
        Console.WriteLine(list.Count & "|" & (list.Capacity >= 50))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0|True"]);
}

#[test]
fn test_vb_list_remove_all_struct_items() {
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
        Dim pts As New List(Of Point) From {New Point(0, 0), New Point(1, 2), New Point(0, 5)}
        pts.RemoveAll(Function(p) p.X = 0)
        Console.WriteLine(pts.Count & ":" & pts(0).X & "," & pts(0).Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:1,2"]);
}

#[test]
fn test_vb_list_remove_all_nullable_types() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Nullable(Of Integer)) From {1, Nothing, 3, Nothing, 5}
        list.RemoveAll(Function(item) Not item.HasValue)
        Console.WriteLine(list.Count & "|" & list(0).Value & "," & list(1).Value & "," & list(2).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3|1,3,5"]);
}

#[test]
fn test_vb_list_remove_all_side_effects_tracking() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim evaluated As Integer = 0
        Dim list As New List(Of Integer) From {10, 20, 30, 40}
        list.RemoveAll(Function(n)
            evaluated += 1
            Return n > 25
        End Function)
        Console.WriteLine(evaluated & "|" & String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4|10,20"]);
}

#[test]
fn test_vb_list_remove_all_enum_values() {
    let src = r#"
Imports System.Collections.Generic

Enum Priority
    Low
    High
End Enum

Module Program
    Sub Main()
        Dim list As New List(Of Priority) From {Priority.Low, Priority.High, Priority.Low}
        list.RemoveAll(Function(p) p = Priority.Low)
        Console.WriteLine(list.Count & ":" & list(0).ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:High"]);
}

#[test]
fn test_vb_list_remove_all_address_of_predicate() {
    let src = r#"
Imports System.Collections.Generic

Module Filters
    Public Function IsNegative(n As Integer) As Boolean
        Return n < 0
    End Function
End Module

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {10, -5, 20, -15, 30}
        list.RemoveAll(AddressOf Filters.IsNegative)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_list_remove_all_case_insensitive_strings() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"Foo", "BAR", "foo", "Baz"}
        list.RemoveAll(Function(s) s.Equals("FOO", StringComparison.OrdinalIgnoreCase))
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["BAR,Baz"]);
}

#[test]
fn test_vb_list_remove_all_even_indices_simulation() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim idx As Integer = 0
        Dim list As New List(Of String) From {"A", "B", "C", "D", "E"}
        list.RemoveAll(Function(item)
            Dim isEvenIndex As Boolean = (idx Mod 2 = 0)
            idx += 1
            Return isEvenIndex
        End Function)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B,D"]);
}

#[test]
fn test_vb_list_remove_all_empty_list() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim emptyList As New List(Of Double)()
        Dim count As Integer = emptyList.RemoveAll(Function(d) d > 0)
        Console.WriteLine(count & "|" & emptyList.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0|0"]);
}

#[test]
fn test_vb_list_remove_all_tuple_elements() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim pairs As New List(Of (Key As String, Val As Integer)) From {
            ("A", 1),
            ("B", 0),
            ("C", 2)
        }
        pairs.RemoveAll(Function(p) p.Val = 0)
        Console.WriteLine(pairs.Count & ":" & pairs(0).Key & "," & pairs(1).Key)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2:A,C"]);
}
