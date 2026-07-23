use super::helpers::run_vb;

#[test]
fn queue_basic_operations() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim queue As New Queue(Of Integer)()
        queue.Enqueue(1)
        queue.Enqueue(2)
        queue.Enqueue(3)
        Console.WriteLine(queue.Count)
        Console.WriteLine(queue.Dequeue())
        Console.WriteLine(queue.Peek())
        Console.WriteLine(queue.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "1", "2", "2"]);
}

#[test]
fn queue_contains_and_copyto() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim queue As New Queue(Of String)()
        queue.Enqueue("a")
        queue.Enqueue("b")
        queue.Enqueue("c")
        Console.WriteLine(queue.Contains("b"))
        Dim target(2) As String
        queue.CopyTo(target, 0)
        Console.WriteLine(target(0))
        Console.WriteLine(target(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "a", "c"]);
}

#[test]
fn stack_basic_lifo_operations() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim stack As New Stack(Of Integer)()
        stack.Push(10)
        stack.Push(20)
        Console.WriteLine(stack.Peek())
        Console.WriteLine(stack.Pop())
        Console.WriteLine(stack.Pop())
        Console.WriteLine(stack.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["20", "20", "10", "0"]);
}

#[test]
fn stack_clear_empties_structure() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim stack As New Stack(Of Integer)()
        stack.Push(1)
        stack.Push(2)
        stack.Clear()
        Console.WriteLine(stack.Count = 0)
        Console.WriteLine(stack.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "0"]);
}

#[test]
fn stack_to_array_preserves_order_reversed() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim stack As New Stack(Of Integer)()
        stack.Push(1)
        stack.Push(2)
        stack.Push(3)
        Dim items() As Integer = stack.ToArray()
        Console.WriteLine(items.Length)
        Console.WriteLine(items(0))
        Console.WriteLine(items(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "3", "1"]);
}

#[test]
fn stack_trims_and_iterates_from_top() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim stack As New Stack(Of String)()
        stack.Push("first")
        stack.Push("second")
        stack.Push("third")
        Dim output As String = ""
        For Each value As String In stack
            output &= value & ","
        Next
        Console.WriteLine(output)
        Console.WriteLine(stack.Count = 3)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["third,second,first,", "True"]);
}

#[test]
fn queue_trydequeue_pattern() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim queue As New Queue(Of String)()
        queue.Enqueue("x")
        Dim first As String = queue.Dequeue()
        Console.WriteLine(first)
        Try
            queue.Dequeue()
            Console.WriteLine("extra")
        Catch ex As InvalidOperationException
            Console.WriteLine("empty")
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["x", "empty"]);
}

#[test]
fn stack_trypeek_pattern() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim stack As New Stack(Of String)()
        Dim noValue As String = "empty"
        Try
            noValue = stack.Peek()
        Catch ex As InvalidOperationException
            noValue = "empty"
        End Try
        Console.WriteLine(noValue = "empty")
        stack.Push("head")
        Console.WriteLine(stack.Peek())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "head"]);
}
