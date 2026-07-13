use super::helpers::run_vb;

#[test]
fn multiline_lambda_no_args() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim act As Action = Sub()
                                Console.WriteLine("Action Executed")
                            End Sub
        act()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Action Executed"]);
}
