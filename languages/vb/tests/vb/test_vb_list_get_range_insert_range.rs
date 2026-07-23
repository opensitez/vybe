use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: List(Of T) GetRange, InsertRange & AddRange Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_list_add_range_enumerable() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2}
        Dim extra As Integer() = {3, 4, 5}
        list.AddRange(extra)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,4,5"]);
}

#[test]
fn test_vb_list_insert_range_at_beginning() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"C", "D"}
        Dim prefix As String() = {"A", "B"}
        list.InsertRange(0, prefix)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A,B,C,D"]);
}

#[test]
fn test_vb_list_insert_range_in_middle() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"A", "D"}
        Dim middle As String() = {"B", "C"}
        list.InsertRange(1, middle)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A,B,C,D"]);
}

#[test]
fn test_vb_list_insert_range_at_end() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {10, 20}
        Dim tail As Integer() = {30, 40}
        list.InsertRange(list.Count, tail)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30,40"]);
}

#[test]
fn test_vb_list_get_range_sublist_extraction() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"A", "B", "C", "D", "E"}
        Dim subList As List(Of String) = list.GetRange(1, 3)
        Console.WriteLine(String.Join(",", subList))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B,C,D"]);
}

#[test]
fn test_vb_list_get_range_independence_from_original() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim original As New List(Of Integer) From {1, 2, 3}
        Dim subList As List(Of Integer) = original.GetRange(0, 2)
        subList(0) = 99
        Console.WriteLine(original(0) & ":" & subList(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:99"]);
}

#[test]
fn test_vb_list_get_range_zero_length() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {10, 20, 30}
        Dim emptySub As List(Of Integer) = list.GetRange(1, 0)
        Console.WriteLine(emptySub.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_list_insert_range_empty_collection() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2}
        Dim empty As Integer() = {}
        list.InsertRange(1, empty)
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2"]);
}

#[test]
fn test_vb_list_add_range_from_another_list() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim srcList As New List(Of Double) From {1.1, 2.2}
        Dim destList As New List(Of Double) From {0.0}
        destList.AddRange(srcList)
        Console.WriteLine(String.Join(";", destList))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0;1.1;2.2"]);
}

#[test]
fn test_vb_list_insert_single_item_at_indices() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"B"}
        list.Insert(0, "A")
        list.Insert(2, "C")
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A,B,C"]);
}

#[test]
fn test_vb_list_as_read_only_wrapper() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {10, 20, 30}
        Dim ro As ReadOnlyCollection(Of Integer) = list.AsReadOnly()
        Console.WriteLine(ro.Count & ":" & ro(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3:10"]);
}

#[test]
fn test_vb_list_as_read_only_reflects_underlying_mutations() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"One"}
        Dim ro As ReadOnlyCollection(Of String) = list.AsReadOnly()
        list.Add("Two")
        Console.WriteLine(ro.Count & ":" & ro(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2:Two"]);
}

#[test]
fn test_vb_list_copy_to_array_destination() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {10, 20, 30}
        Dim target(4) As Integer
        list.CopyTo(target, 1) ' Copy starting at target index 1
        Console.WriteLine(String.Join(",", target))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0,10,20,30,0"]);
}

#[test]
fn test_vb_list_copy_to_array_subset() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {10, 20, 30, 40, 50}
        Dim target(1) As Integer
        list.CopyTo(1, target, 0, 2) ' Copy 2 elements starting from list index 1
        Console.WriteLine(String.Join(",", target))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20,30"]);
}

#[test]
fn test_vb_list_to_array_creates_new_array() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"A", "B"}
        Dim arr As String() = list.ToArray()
        Console.WriteLine(arr.Length & ":" & arr(0) & "," & arr(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2:A,B"]);
}

#[test]
fn test_vb_list_find_all_returns_new_list() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 3, 4, 5}
        Dim oddList As List(Of Integer) = list.FindAll(Function(n) n Mod 2 <> 0)
        Console.WriteLine(String.Join(",", oddList))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,3,5"]);
}

#[test]
fn test_vb_list_find_index_find_last_index() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"apple", "banana", "apricot", "cherry"}
        Dim firstA As Integer = list.FindIndex(Function(s) s.StartsWith("a"))
        Dim lastA As Integer = list.FindLastIndex(Function(s) s.StartsWith("a"))
        Console.WriteLine(firstA & ":" & lastA)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0:2"]);
}

#[test]
fn test_vb_list_exists_true_for_all() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {2, 4, 6, 8}
        Dim hasEight As Boolean = list.Exists(Function(n) n = 8)
        Dim allEven As Boolean = list.TrueForAll(Function(n) n Mod 2 = 0)
        Console.WriteLine(hasEight & "|" & allEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_list_for_each_action_invocation() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim sum As Integer = 0
        Dim list As New List(Of Integer) From {10, 20, 30}
        list.ForEach(Sub(n) sum += n)
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["60"]);
}

#[test]
fn test_vb_list_reverse_in_place_and_range() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 3, 4, 5}
        list.Reverse(1, 3) ' Reverse elements from index 1 (len 3): 2,3,4 -> 4,3,2
        Console.WriteLine(String.Join(",", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,4,3,2,5"]);
}
