use super::helpers::run_vb;

#[test]
fn async_await_syntax() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Async Function GetDataAsync() As Task(Of Integer)
        Await Task.Delay(1)
        Return 42
    End Function

    Sub Main()
        ' Using Wait() for console app simplicity, testing Async/Await compilation
        Dim task = GetDataAsync()
        task.Wait()
        Console.WriteLine(task.Result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}
