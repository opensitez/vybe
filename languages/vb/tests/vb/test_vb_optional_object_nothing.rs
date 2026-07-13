use super::helpers::run_vb;

#[test]
fn optional_object_nothing() {
    let out = run_vb(
        r#"
Module M
    ' Optional Object parameter defaulting to Nothing
    Sub DoWork(Optional obj As Object = Nothing)
        Console.WriteLine(obj Is Nothing)
    End Sub

    Sub Main()
        DoWork()
        DoWork(New Object())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}
