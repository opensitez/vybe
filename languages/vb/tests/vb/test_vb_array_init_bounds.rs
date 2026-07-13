use super::helpers::run_vb;

#[test]
fn array_initialization_1d() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr1 As Integer() = {1, 2, 3}
        Dim arr2() As Integer = {4, 5, 6}
        Console.WriteLine(arr1(0) + arr2(0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn array_initialization_2d() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr(,) As Integer = {{1, 2}, {3, 4}}
        Console.WriteLine(arr(1, 1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn array_bounds() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim arr(5) As Integer
        Console.WriteLine(arr.Length) ' 6 elements!
        Console.WriteLine(arr.GetUpperBound(0)) ' 5
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["6", "5"]);
}
