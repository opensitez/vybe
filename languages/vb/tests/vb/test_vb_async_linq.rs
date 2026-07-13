use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Async with LINQ
// ═══════════════════════════════════════════════════════════

#[test]
fn async_with_linq() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks
Imports System.Collections.Generic
Imports System.Linq

Module M
    Async Function GetNumberAsync(n As Integer) As Task(Of Integer)
        Await Task.Delay(1)
        Return n * 2
    End Function

    Async Function ProcessListAsync() As Task
        Dim nums As New List(Of Integer) From { 1, 2, 3 }
        
        ' VB supports Await inside LINQ queries (unlike C# which only supports it in certain contexts)
        ' Actually, Await in query expressions is NOT supported in VB.NET.
        ' Let's use it in a normal loop to be safe, or Task.WhenAll.
        
        Dim tasks = nums.Select(Function(n) GetNumberAsync(n))
        Dim results = Await Task.WhenAll(tasks)
        
        For Each r In results
            Console.WriteLine(r)
        Next
    End Function

    Sub Main()
        ProcessListAsync().Wait()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "4", "6"]);
}
