use super::helpers::run_vb;

#[test]
fn list_initializer_has_expected_count_and_order() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {1, 2, 3}
        Console.WriteLine(values.Count)
        Console.WriteLine(values(0))
        Console.WriteLine(values(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "1", "3"]);
}

#[test]
fn list_add_and_remove_affects_count_and_contains() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer)()
        values.Add(9)
        values.Add(10)
        Console.WriteLine(values.Count)
        Console.WriteLine(values.Contains(10))
        values.Remove(10)
        Console.WriteLine(values.Count)
        Console.WriteLine(values.Contains(10))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "True", "1", "False"]);
}

#[test]
fn list_insert_pushes_item_into_middle() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of String) From {"a", "c"}
        values.Insert(1, "b")
        Console.WriteLine(values(0))
        Console.WriteLine(values(1))
        Console.WriteLine(values(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn list_remove_at_removes_selected_element() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {5, 6, 7, 8}
        values.RemoveAt(1)
        Console.WriteLine(values.Count)
        Console.WriteLine(values(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "7"]);
}

#[test]
fn list_remove_at_shifts_tail_forward() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {1, 2, 3, 4}
        values.RemoveAt(0)
        Console.WriteLine(values(0))
        Console.WriteLine(values(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn list_index_of_detects_position_and_absent_key() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {7, 8, 9}
        Console.WriteLine(values.IndexOf(8))
        Console.WriteLine(values.IndexOf(99))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "-1"]);
}

#[test]
fn list_foreach_sums_values() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {1, 2, 3}
        Dim total As Integer = 0
        For Each value As Integer In values
            total += value
        Next
        Console.WriteLine(total)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["6"]);
}

#[test]
fn list_sort_reorders_values() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {5, 1, 4, 2}
        values.Sort()
        Console.WriteLine(values(0))
        Console.WriteLine(values(3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn list_reverse_inverts_order() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {1, 2, 3}
        values.Reverse()
        For Each value As Integer In values
            Console.WriteLine(value)
        Next
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn list_clear_empties_items() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {1, 2, 3}
        values.Clear()
        Console.WriteLine(values.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn list_find_returns_first_matching_predicate() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {2, 5, 7, 9}
        Dim hit As Integer = values.Find(Function(value As Integer) value > 6)
        Console.WriteLine(hit)
        Dim noneHit As Integer = values.Find(Function(value As Integer) value > 20)
        Console.WriteLine(noneHit = 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["7", "True"]);
}

#[test]
fn list_find_all_counts_matches() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {1, 2, 3, 4}
        Dim evens As List(Of Integer) = values.FindAll(Function(value As Integer) value Mod 2 = 0)
        Console.WriteLine(evens.Count)
        Console.WriteLine(evens(0))
        Console.WriteLine(evens(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "2", "4"]);
}

#[test]
fn list_exists_returns_predicate_match() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {3, 5, 8}
        Console.WriteLine(values.Exists(Function(v As Integer) v > 5))
        Console.WriteLine(values.Exists(Function(v As Integer) v < 0))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn list_binarysearch_finds_in_sorted_list() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {1, 3, 5, 7}
        Console.WriteLine(values.BinarySearch(5))
        Console.WriteLine(values.BinarySearch(6))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "-4"]);
}
