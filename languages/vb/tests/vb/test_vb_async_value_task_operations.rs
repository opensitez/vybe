use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ValueTask(Of T) Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_async_value_task_synchronous_completion() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Function GetCachedValueAsync(cached As Boolean) As ValueTask(Of Integer)
        If cached Then
            Return New ValueTask(Of Integer)(100)
        End If
        Return New ValueTask(Of Integer)(ComputeAsync())
    End Function

    Async Function ComputeAsync() As Task(Of Integer)
        Await Task.Delay(10)
        Return 200
    End Function

    Async Function RunAsync() As Task
        Dim v1 As Integer = Await GetCachedValueAsync(True)
        Dim v2 As Integer = Await GetCachedValueAsync(False)
        Console.WriteLine(v1 & ":" & v2)
    End Function

    Sub Main()
        RunAsync().Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100:200"]);
}
