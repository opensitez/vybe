use super::helpers::run_vb;

#[test]
fn try_catch_finally() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim state As String = "Start"
        Try
            state = "Try"
            Throw New System.Exception("Fail")
        Catch ex As System.Exception
            state = "Catch"
        Finally
            state = "Finally"
        End Try
        Console.WriteLine(state)
        
        ' Catch without exception variable
        Try
            Throw New System.Exception()
        Catch
            Console.WriteLine("Caught")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Finally", "Caught"]);
}
