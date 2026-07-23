use super::helpers::run_vb;

#[test]
fn hashset_union_includes_all_elements() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim left As New HashSet(Of Integer)()
        left.Add(1)
        left.Add(2)

        Dim right As New HashSet(Of Integer)()
        right.Add(2)
        right.Add(3)

        left.UnionWith(right)
        Console.WriteLine(left.Count)
        Console.WriteLine(left.Contains(3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "True"]);
}

#[test]
fn hashset_intersect_keeps_common() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim baseSet As New HashSet(Of Integer)()
        baseSet.Add(1)
        baseSet.Add(2)
        baseSet.Add(3)

        Dim mask As New HashSet(Of Integer)()
        mask.Add(2)
        mask.Add(3)
        mask.Add(4)

        baseSet.IntersectWith(mask)
        Console.WriteLine(baseSet.Count)
        Console.WriteLine(baseSet.Contains(1))
        Console.WriteLine(baseSet.Contains(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "False", "True"]);
}

#[test]
fn hashset_except_removes_overlap() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim baseSet As New HashSet(Of Integer)()
        baseSet.Add(1)
        baseSet.Add(2)
        baseSet.Add(3)

        Dim toRemove As New HashSet(Of Integer)()
        toRemove.Add(2)

        baseSet.ExceptWith(toRemove)
        Console.WriteLine(baseSet.Count)
        Console.WriteLine(baseSet.Contains(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "False"]);
}

#[test]
fn hashset_symmetric_except_with_flips_to_symmetric_difference() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim left As New HashSet(Of Integer)()
        left.Add(1)
        left.Add(2)
        left.Add(3)

        Dim right As New HashSet(Of Integer)()
        right.Add(2)
        right.Add(4)

        left.SymmetricExceptWith(right)
        Console.WriteLine(left.Count)
        Console.WriteLine(left.Contains(1))
        Console.WriteLine(left.Contains(2))
        Console.WriteLine(left.Contains(4))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "True", "False", "True"]);
}

#[test]
fn hashset_set_equals_and_subset_checks() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim a As New HashSet(Of Integer)()
        a.Add(1)
        a.Add(2)

        Dim b As New HashSet(Of Integer)()
        b.Add(1)
        b.Add(2)

        Dim c As New HashSet(Of Integer)()
        c.Add(1)

        Console.WriteLine(a.SetEquals(b))
        Console.WriteLine(c.IsSubsetOf(a))
        Console.WriteLine(a.IsSupersetOf(c))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn hashset_overlaps_reports_commonality() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim a As New HashSet(Of Integer)()
        a.Add(1)
        a.Add(2)

        Dim b As New HashSet(Of Integer)()
        b.Add(2)
        b.Add(9)

        Console.WriteLine(a.Overlaps(b))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
