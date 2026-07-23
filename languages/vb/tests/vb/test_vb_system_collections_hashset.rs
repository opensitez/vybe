use super::helpers::run_vb;

#[test]
fn system_collections_hashset_union_and_superset() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim left As New HashSet(Of Integer)()
        left.Add(1)
        left.Add(3)

        Dim right As New HashSet(Of Integer)()
        right.Add(3)
        right.Add(4)

        left.UnionWith(right)
        Console.WriteLine(left.Count)
        Console.WriteLine(left.Contains(4))
        Console.WriteLine(left.IsSupersetOf(right))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "True", "True"]);
}

#[test]
fn system_collections_hashset_intersect_and_except() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim baseSet As New HashSet(Of Integer)()
        baseSet.Add(1)
        baseSet.Add(2)
        baseSet.Add(3)

        Dim intersectWith As New HashSet(Of Integer)()
        intersectWith.Add(2)
        intersectWith.Add(3)
        intersectWith.Add(4)

        baseSet.IntersectWith(intersectWith)
        Console.WriteLine(baseSet.Count)
        Console.WriteLine(baseSet.Contains(2))

        baseSet.ExceptWith(intersectWith)
        Console.WriteLine(baseSet.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "True", "0"]);
}

#[test]
fn system_collections_hashset_set_relations() {
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

        Console.WriteLine(a.SetEquals(b))
        Console.WriteLine(a.Overlaps(b))
        Console.WriteLine(a.IsProperSubsetOf(b))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "False"]);
}
