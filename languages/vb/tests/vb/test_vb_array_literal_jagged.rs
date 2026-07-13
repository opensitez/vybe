use super::helpers::run_vb;

#[test]
fn array_literal_jagged() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Jagged array literal inference
        Dim arr = {({1, 2}), ({3, 4, 5})}
        
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr(1).Length)
        Console.WriteLine(arr(1)(2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "3", "5"]);
}
