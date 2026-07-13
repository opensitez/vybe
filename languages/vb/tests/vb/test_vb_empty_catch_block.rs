use super::helpers::run_vb;

#[test]
fn empty_catch_block() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Dim i = 1 \ 0
        Catch
            ' Empty catch block to swallow exceptions
        End Try
        Console.WriteLine("Survived")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Survived"]);
}
