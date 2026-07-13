use super::helpers::run_vb;

#[test]
fn await_in_catch() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Async Function LogErrorAsync() As Task
        Console.WriteLine("Logged")
    End Function

    Async Function TestAsync() As Task
        Try
            Throw New System.Exception()
        Catch ex As System.Exception
            ' Await inside Catch (added in VB 14)
            Await LogErrorAsync()
        End Try
    End Function

    Sub Main()
        TestAsync().Wait()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Logged"]);
}
