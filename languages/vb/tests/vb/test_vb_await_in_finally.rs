use super::helpers::run_vb;

#[test]
fn await_in_finally() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Async Function CleanupAsync() As Task
        Console.WriteLine("Cleaned")
    End Function

    Async Function TestAsync() As Task
        Try
            ' do nothing
        Finally
            ' Await inside Finally (added in VB 14)
            Await CleanupAsync()
        End Try
    End Function

    Sub Main()
        TestAsync().Wait()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Cleaned"]);
}
