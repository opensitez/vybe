use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Async Methods in Interfaces
// ═══════════════════════════════════════════════════════════

#[test]
fn async_interfaces() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Interface IWorker
    Function DoWorkAsync() As Task(Of Integer)
End Interface

Class Worker
    Implements IWorker
    
    ' The Async modifier goes on the implementation, not the interface
    Public Async Function DoWorkAsync() As Task(Of Integer) Implements IWorker.DoWorkAsync
        Await Task.Delay(1)
        Return 42
    End Function
End Class

Module M
    Sub Main()
        Dim w As IWorker = New Worker()
        Dim t As Task(Of Integer) = w.DoWorkAsync()
        t.Wait()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}
