use super::helpers::run_vb;

#[test]
fn string_interpolation_expr() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x = 10
        Dim y = 20
        
        ' String interpolation with expressions
        Dim s = $"Result is {x + y}"
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Result is 30"]);
}
