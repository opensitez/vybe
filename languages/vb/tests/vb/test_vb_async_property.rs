use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Async in Properties (Not natively supported, testing workarounds)
// ═══════════════════════════════════════════════════════════

#[test]
fn async_property_workaround() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Class DataService
    ' Properties cannot be Async directly. 
    ' But they can return a Task(Of T).
    Public ReadOnly Property DataAsync As Task(Of String)
        Get
            Return FetchDataAsync()
        End Get
    End Property
    
    Private Async Function FetchDataAsync() As Task(Of String)
        Await Task.Delay(1)
        Return "Async Data"
    End Function
End Class

Module M
    Sub Main()
        Dim ds As New DataService()
        ' We synchronously wait for the task in Main
        Dim result = ds.DataAsync.Result
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Async Data"]);
}
