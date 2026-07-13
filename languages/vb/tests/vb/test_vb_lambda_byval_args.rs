use super::helpers::run_vb;

#[test]
fn lambda_byval_args() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Explicit ByVal in lambda arguments
        Dim act As Action(Of Integer) = Sub(ByVal x As Integer)
                                            x += 10
                                        End Sub
        Dim val = 5
        act(val)
        Console.WriteLine(val) ' Should still be 5
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5"]);
}
