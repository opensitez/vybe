use super::helpers::run_vb;

#[test]
fn explicit_implements_different_name() {
    let out = run_vb(
        r#"
Interface IWorker
    Sub DoWork()
End Interface

Class Worker
    Implements IWorker
    
    ' In VB.NET, the method name doesn't have to match the interface method name,
    ' the Implements clause explicitly links them.
    Public Sub PerformTask() Implements IWorker.DoWork
        Console.WriteLine("Working")
    End Sub
End Class

Module M
    Sub Main()
        Dim w As IWorker = New Worker()
        w.DoWork()
        
        Dim c As Worker = New Worker()
        c.PerformTask()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Working", "Working"]);
}
