use super::helpers::run_vb;

#[test]
fn array_literal_bounds() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Array literal bounds inference for multi-dimensional arrays
        Dim arr(,) = {{1, 2}, {3, 4}, {5, 6}}
        
        Console.WriteLine(arr.GetLength(0))
        Console.WriteLine(arr.GetLength(1))
        Console.WriteLine(arr(2, 1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "2", "6"]);
}
