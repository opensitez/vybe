use super::helpers::run_vb;

#[test]
fn delegate_type_inference() {
    let out = run_vb(
        r#"
Module M
    Sub DoWork(action As Action)
        action()
    End Sub

    Sub Main()
        ' Delegate inference from Lambda
        DoWork(Sub() Console.WriteLine("Inferred"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Inferred"]);
}
