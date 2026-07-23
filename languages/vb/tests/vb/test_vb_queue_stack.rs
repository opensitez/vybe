use super::helpers::run_vb;

#[test]
fn queue_enqueue_dequeue_fifo_order() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim q As New Queue(Of Integer)()
        q.Enqueue(10)
        q.Enqueue(20)
        Console.WriteLine(q.Dequeue())
        Console.WriteLine(q.Dequeue())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn queue_peek_and_count_after_enqueue() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim q As New Queue(Of Integer)()
        q.Enqueue(1)
        q.Enqueue(2)
        Console.WriteLine(q.Peek())
        Console.WriteLine(q.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn queue_contains_reports_membership() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim q As New Queue(Of Integer)()
        q.Enqueue(5)
        q.Enqueue(6)
        Console.WriteLine(q.Contains(6))
        Console.WriteLine(q.Contains(9))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn queue_clear_empties_all_items() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim q As New Queue(Of Integer)()
        q.Enqueue(1)
        q.Clear()
        Console.WriteLine(q.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn queue_copy_from_array_roundtrip() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim source As Integer() = {9, 8, 7}
        Dim q As New Queue(Of Integer)(source)
        Console.WriteLine(q.Count)
        Console.WriteLine(q.Dequeue())
        Console.WriteLine(q.Dequeue())
        Console.WriteLine(q.Dequeue())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "9", "8", "7"]);
}

#[test]
fn stack_push_pop_lifo_order() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim s As New Stack(Of Integer)()
        s.Push(1)
        s.Push(2)
        s.Push(3)
        Console.WriteLine(s.Pop())
        Console.WriteLine(s.Pop())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn stack_peek_reports_top_without_removing() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim s As New Stack(Of Integer)()
        s.Push(1)
        s.Push(2)
        Console.WriteLine(s.Peek())
        Console.WriteLine(s.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn stack_contains_reports_membership() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim s As New Stack(Of Integer)()
        s.Push(5)
        s.Push(6)
        Console.WriteLine(s.Contains(6))
        Console.WriteLine(s.Contains(9))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn stack_clear_empties_all_items() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim s As New Stack(Of Integer)()
        s.Push(1)
        s.Clear()
        Console.WriteLine(s.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn stack_pop_from_empty_throws_exception() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim s As New Stack(Of Integer)()
        Try
            s.Pop()
            Console.WriteLine("NoError")
        Catch ex As InvalidOperationException
            Console.WriteLine("Error")
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["Error"]);
}
