use super::helpers::run_vb;

#[test]
fn hashset_add_remove_count() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim set As New HashSet(Of Integer)()
        Console.WriteLine(set.Add(1))
        Console.WriteLine(set.Add(1))
        set.Remove(1)
        Console.WriteLine(set.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False", "0"]);
}

#[test]
fn hashset_contains_behaviour() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim set As New HashSet(Of String)()
        set.Add("a")
        set.Add("b")
        Console.WriteLine(set.Contains("a"))
        Console.WriteLine(set.Contains("c"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn hashset_set_equal_contract() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim left As New HashSet(Of Integer) From {1, 2, 3}
        Dim right As New HashSet(Of Integer) From {3, 2, 1}
        Console.WriteLine(left.SetEquals(right))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn hashset_union_and_intersection() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim left As New HashSet(Of Integer) From {1, 2}
        Dim right As New HashSet(Of Integer) From {2, 3}
        left.UnionWith(right)
        Console.WriteLine(left.Count)
        Dim intersection As New HashSet(Of Integer)({1, 2, 3})
        intersection.IntersectWith(New HashSet(Of Integer)({2, 3, 4}))
        Console.WriteLine(intersection.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn hashset_except_and_symmetric_difference() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim left As New HashSet(Of Integer) From {1, 2, 3, 4}
        left.ExceptWith(New HashSet(Of Integer)({3, 4}))
        Console.WriteLine(left.Count)
        Dim a As New HashSet(Of Integer) From {1, 2, 5}
        Dim b As New HashSet(Of Integer) From {2, 3}
        a.SymmetricExceptWith(b)
        Console.WriteLine(a.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn hashset_overlaps_subset_is_subset() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim big As New HashSet(Of Integer) From {1, 2, 3, 4}
        Dim small As New HashSet(Of Integer) From {2, 3}
        Console.WriteLine(big.Overlaps(small))
        Console.WriteLine(small.IsSubsetOf(big))
        Console.WriteLine(big.IsSupersetOf(small))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn hashset_ensure_non_references_unique() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim set As New HashSet(Of String)()
        set.Add("same")
        set.Add(New String("same".ToCharArray()))
        Console.WriteLine(set.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1"]);
}
