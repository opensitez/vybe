use super::helpers::run_vb;

#[test]
fn throw_without_args_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Try
                Throw New System.Exception("Original")
            Catch
                ' Throw without args re-throws the current exception
                Throw
            End Try
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Original"]);
}
