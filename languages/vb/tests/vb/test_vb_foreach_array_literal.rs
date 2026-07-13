use super::helpers::run_vb;

#[test]
fn foreach_array_literal() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' For Each with implicit array literal
        For Each x In {10, 20, 30}
            Console.WriteLine(x)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}
