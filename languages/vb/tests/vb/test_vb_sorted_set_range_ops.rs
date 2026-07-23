use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: SortedSet(Of T) Range & Set Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_sorted_set_sorted_order() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ss As New SortedSet(Of Integer) From {5, 1, 9, 3, 7}
        Console.WriteLine(String.Join(",", ss))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,3,5,7,9"]);
}

#[test]
fn test_vb_sorted_set_min_max_properties() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ss As New SortedSet(Of Integer) From {40, 10, 50, 20}
        Console.WriteLine(ss.Min)
        Console.WriteLine(ss.Max)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10", "50"]);
}

#[test]
fn test_vb_sorted_set_get_view_between_range() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ss As New SortedSet(Of Integer) From {10, 20, 30, 40, 50, 60}
        Dim view As SortedSet(Of Integer) = ss.GetViewBetween(20, 50)
        Console.WriteLine(String.Join(",", view))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20,30,40,50"]);
}

#[test]
fn test_vb_sorted_set_reverse_view() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ss As New SortedSet(Of String) From {"A", "B", "C", "D"}
        Dim rev As IEnumerable(Of String) = ss.Reverse()
        Console.WriteLine(String.Join(",", rev))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["D,C,B,A"]);
}

#[test]
fn test_vb_sorted_set_custom_comparer() {
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
        Dim ss As New SortedSet(Of Integer)(New DescendingIntComparer()) From {10, 30, 20}
        Console.WriteLine(String.Join(",", ss))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30,20,10"]);
}

#[test]
fn test_vb_sorted_set_union_intersect() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim s1 As New SortedSet(Of Integer) From {1, 2, 3}
        Dim s2 As New SortedSet(Of Integer) From {3, 4, 5}
        s1.UnionWith(s2)
        Console.WriteLine(String.Join(",", s1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,4,5"]);
}

#[test]
fn test_vb_sorted_set_remove_where() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ss As New SortedSet(Of Integer) From {1, 2, 3, 4, 5}
        ss.RemoveWhere(Function(x) x Mod 2 <> 0)
        Console.WriteLine(String.Join(",", ss))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2,4"]);
}
