use super::helpers::run_vb;

#[test]
fn linq_orderby_default_orderings() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As Integer() = {5, 1, 3, 2}
        Dim ordered = values.OrderBy(Function(v) v).ToArray()
        Dim descending = values.OrderByDescending(Function(v) v).ToArray()

        Console.WriteLine(String.Join(",", ordered))
        Console.WriteLine(String.Join(",", descending))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1,2,3,5", "5,3,2,1"]);
}

#[test]
fn linq_ordering_secondary_by_key() {
    let out = run_vb(
        r#"
Class Item
    Public Name As String
    Public Score As Integer

    Public Sub New(name As String, score As Integer)
        Me.Name = name
        Me.Score = score
    End Sub
End Class

Module M
    Sub Main()
        Dim data = {
            New Item("c", 1),
            New Item("a", 2),
            New Item("b", 2),
            New Item("a", 1)
        }

        Dim sorted = data.OrderBy(Function(i) i.Score).ThenBy(Function(i) i.Name)
        Dim firstName As String = sorted(0).Name
        Dim lastName As String = sorted.Last().Name

        Console.WriteLine(sorted.First().Score)
        Console.WriteLine(firstName)
        Console.WriteLine(lastName)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "a", "c"]);
}

#[test]
fn linq_ordering_reversible() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim values As String() = {"b", "a", "c"}
        Dim ascending = values.OrderBy(Function(v) v)
        Dim descending = ascending.Reverse()

        Console.WriteLine(ascending.First())
        Console.WriteLine(descending.First())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["a", "c"]);
}
