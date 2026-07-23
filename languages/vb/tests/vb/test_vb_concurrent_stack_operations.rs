use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ConcurrentStack(Of T) Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_concurrent_stack_push_try_pop() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cs As New ConcurrentStack(Of Integer)()
        cs.Push(10)
        cs.Push(20)
        Dim item As Integer
        Dim ok1 As Boolean = cs.TryPop(item)
        Console.WriteLine(ok1)
        Console.WriteLine(item)
        Dim ok2 As Boolean = cs.TryPop(item)
        Console.WriteLine(ok2)
        Console.WriteLine(item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "20", "True", "10"]);
}

#[test]
fn test_vb_concurrent_stack_push_range_try_pop_range() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cs As New ConcurrentStack(Of Integer)()
        Dim items As Integer() = {1, 2, 3, 4}
        cs.PushRange(items)
        Dim popped(1) As Integer
        Dim count As Integer = cs.TryPopRange(popped)
        Console.WriteLine(count)
        Console.WriteLine(String.Join(",", popped))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "4,3"]);
}

#[test]
fn test_vb_concurrent_stack_try_peek() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cs As New ConcurrentStack(Of String)()
        cs.Push("Top")
        Dim topVal As String = Nothing
        Dim ok As Boolean = cs.TryPeek(topVal)
        Console.WriteLine(ok)
        Console.WriteLine(topVal)
        Console.WriteLine(cs.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "Top", "1"]);
}

#[test]
fn test_vb_concurrent_stack_clear() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cs As New ConcurrentStack(Of Integer)()
        cs.Push(100)
        cs.Clear()
        Console.WriteLine(cs.IsEmpty)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
