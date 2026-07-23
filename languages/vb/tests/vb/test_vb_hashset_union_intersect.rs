use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: HashSet(Of T) Set Algebra & Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_hashset_add_duplicate_prevention() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim hs As New HashSet(Of Integer)()
        Dim added1 As Boolean = hs.Add(10)
        Dim added2 As Boolean = hs.Add(10)
        Console.WriteLine(added1)
        Console.WriteLine(added2)
        Console.WriteLine(hs.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False", "1"]);
}

#[test]
fn test_vb_hashset_union_with() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set1 As New HashSet(Of Integer) From {1, 2, 3}
        Dim set2 As New HashSet(Of Integer) From {3, 4, 5}
        set1.UnionWith(set2)
        Console.WriteLine(set1.Count)
        Console.WriteLine(String.Join(",", set1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5", "1,2,3,4,5"]);
}

#[test]
fn test_vb_hashset_intersect_with() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set1 As New HashSet(Of Integer) From {1, 2, 3, 4}
        Dim set2 As New HashSet(Of Integer) From {3, 4, 5, 6}
        set1.IntersectWith(set2)
        Console.WriteLine(String.Join(",", set1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3,4"]);
}

#[test]
fn test_vb_hashset_except_with() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set1 As New HashSet(Of Integer) From {1, 2, 3, 4}
        Dim set2 As New HashSet(Of Integer) From {2, 4}
        set1.ExceptWith(set2)
        Console.WriteLine(String.Join(",", set1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,3"]);
}

#[test]
fn test_vb_hashset_symmetric_except_with() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set1 As New HashSet(Of Integer) From {1, 2, 3}
        Dim set2 As New HashSet(Of Integer) From {2, 3, 4}
        set1.SymmetricExceptWith(set2)
        Console.WriteLine(String.Join(",", set1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,4"]);
}

#[test]
fn test_vb_hashset_is_subset_is_superset() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim subSet As New HashSet(Of Integer) From {1, 2}
        Dim superSet As New HashSet(Of Integer) From {1, 2, 3, 4}
        Console.WriteLine(subSet.IsSubsetOf(superSet))
        Console.WriteLine(superSet.IsSupersetOf(subSet))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_hashset_overlaps_and_set_equals() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set1 As New HashSet(Of Integer) From {1, 2, 3}
        Dim set2 As New HashSet(Of Integer) From {3, 4, 5}
        Dim set3 As New HashSet(Of Integer) From {3, 2, 1}
        Console.WriteLine(set1.Overlaps(set2))
        Console.WriteLine(set1.SetEquals(set3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_hashset_remove_where_predicate() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set1 As New HashSet(Of Integer) From {1, 2, 3, 4, 5, 6}
        Dim removed As Integer = set1.RemoveWhere(Function(x) x Mod 2 = 0)
        Console.WriteLine(removed)
        Console.WriteLine(String.Join(",", set1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "1,3,5"]);
}

#[test]
fn test_vb_hashset_custom_equality_comparer() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim hs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {"abc"}
        Dim added As Boolean = hs.Add("ABC")
        Console.WriteLine(added)
        Console.WriteLine(hs.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "1"]);
}
