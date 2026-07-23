use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Queue(Of T) Operations & Semantics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_queue_enqueue_dequeue_fifo_order() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim q As New Queue(Of Integer)()
        q.Enqueue(10)
        q.Enqueue(20)
        q.Enqueue(30)
        Console.WriteLine(q.Dequeue())
        Console.WriteLine(q.Dequeue())
        Console.WriteLine(q.Dequeue())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10", "20", "30"]);
}

#[test]
fn test_vb_queue_peek_non_destructive() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim q As New Queue(Of String)()
        q.Enqueue("First")
        Console.WriteLine(q.Peek())
        Console.WriteLine(q.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First", "1"]);
}

#[test]
fn test_vb_queue_try_peek_try_dequeue() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim q As New Queue(Of String)()
        Dim frontVal As String = Nothing
        Dim okPeek As Boolean = q.TryPeek(frontVal)
        Dim okDeq As Boolean = q.TryDequeue(frontVal)
        Console.WriteLine(okPeek)
        Console.WriteLine(okDeq)

        q.Enqueue("Item")
        okDeq = q.TryDequeue(frontVal)
        Console.WriteLine(okDeq)
        Console.WriteLine(frontVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "False", "True", "Item"]);
}

#[test]
fn test_vb_queue_contains_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim q As New Queue(Of Integer)()
        q.Enqueue(100)
        q.Enqueue(200)
        Console.WriteLine(q.Contains(100))
        Console.WriteLine(q.Contains(300))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_queue_to_array_order() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim q As New Queue(Of Integer)()
        q.Enqueue(1)
        q.Enqueue(2)
        q.Enqueue(3)
        Dim arr As Integer() = q.ToArray()
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_queue_clear() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim q As New Queue(Of Integer)()
        q.Enqueue(42)
        q.Clear()
        Console.WriteLine(q.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_queue_trim_excess() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim q As New Queue(Of Integer)(100)
        q.Enqueue(1)
        q.TrimExcess()
        Console.WriteLine(q.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}
