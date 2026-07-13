use super::helpers::run_vb;

#[test]
fn linq_where_select() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim nums() = {1, 2, 3, 4, 5, 6}
        Dim q = From n In nums Where n Mod 2 = 0 Select n * 2
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4", "8", "12"]);
}

#[test]
fn linq_order_by() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim names() = {"Charlie", "Alice", "Bob"}
        Dim q = From n In names Order By n Descending Select n
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Charlie", "Bob", "Alice"]);
}

#[test]
fn linq_group_by_comprehensive() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim nums() = {1, 2, 3, 4, 5}
        Dim q = From n In nums Group By IsEven = (n Mod 2 = 0) Into Group Select IsEven, Group
        For Each g In q
            Console.WriteLine(g.IsEven)
            Console.WriteLine(g.Group.Count())
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["False", "3", "True", "2"]);
}

#[test]
fn linq_aggregate_sum() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim nums() = {10, 20, 30}
        Dim sum = Aggregate n In nums Into Sum()
        Console.WriteLine(sum)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn linq_aggregate_average() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim nums() = {10, 20, 30}
        Dim avg = Aggregate n In nums Into Average()
        Console.WriteLine(avg)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn linq_aggregate_max_min() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim nums() = {15, 2, 88, 42}
        Dim mx = Aggregate n In nums Into Max()
        Dim mn = Aggregate n In nums Into Min()
        Console.WriteLine(mx)
        Console.WriteLine(mn)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["88", "2"]);
}

#[test]
fn linq_let_clause() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim nums() = {1, 2, 3}
        Dim q = From n In nums
                Let sq = n * n
                Where sq > 4
                Select sq
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn linq_take_skip() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim nums() = {1, 2, 3, 4, 5}
        Dim q = From n In nums Skip 2 Take 2 Select n
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn linq_take_while() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim nums() = {1, 2, 3, 4, 1, 2}
        Dim q = From n In nums Take While n < 4 Select n
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn linq_skip_while() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim nums() = {1, 2, 3, 4, 1, 2}
        Dim q = From n In nums Skip While n < 4 Select n
        For Each v In q
            Console.WriteLine(v)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4", "1", "2"]);
}
