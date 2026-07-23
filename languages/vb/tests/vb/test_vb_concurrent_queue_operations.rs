use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ConcurrentQueue(Of T) Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_concurrent_queue_enqueue_try_dequeue() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cq As New ConcurrentQueue(Of String)()
        cq.Enqueue("First")
        cq.Enqueue("Second")
        Dim item As String = Nothing
        Dim ok1 As Boolean = cq.TryDequeue(item)
        Console.WriteLine(ok1)
        Console.WriteLine(item)
        Dim ok2 As Boolean = cq.TryDequeue(item)
        Console.WriteLine(ok2)
        Console.WriteLine(item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "First", "True", "Second"]);
}

#[test]
fn test_vb_concurrent_queue_try_peek() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cq As New ConcurrentQueue(Of Integer)()
        cq.Enqueue(100)
        Dim val As Integer
        Dim ok As Boolean = cq.TryPeek(val)
        Console.WriteLine(ok)
        Console.WriteLine(val)
        Console.WriteLine(cq.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "100", "1"]);
}

#[test]
fn test_vb_concurrent_queue_is_empty() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cq As New ConcurrentQueue(Of Double)()
        Console.WriteLine(cq.IsEmpty)
        cq.Enqueue(1.1)
        Console.WriteLine(cq.IsEmpty)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_concurrent_queue_clear() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cq As New ConcurrentQueue(Of Integer)()
        cq.Enqueue(10)
        cq.Enqueue(20)
        cq.Clear()
        Console.WriteLine(cq.IsEmpty)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
