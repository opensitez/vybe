use super::helpers::run_vb;

#[test]
fn singleline_lambda_nothing() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim func As Func(Of Object) = Function() Nothing
        
        Console.WriteLine(func() Is Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
