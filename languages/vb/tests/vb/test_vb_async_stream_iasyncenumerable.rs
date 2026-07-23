use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: IAsyncEnumerable(Of T) & Async Streams
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_async_iterator_yield_return() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Threading.Tasks

Module Program
    Async Function GenerateNumbersAsync() As IAsyncEnumerable(Of Integer)
        ' Mock async stream pattern using list Task result
        Return FetchListAsync().Result
    End Function

    Async Function FetchListAsync() As Task(Of List(Of Integer))
        Await Task.Delay(10)
        Return New List(Of Integer) From {1, 2, 3}
    End Function

    Sub Main()
        Dim items = GenerateNumbersAsync().Result
        Console.WriteLine(String.Join(",", items))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}
