use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Collections.Concurrent.ConcurrentStack LIFO
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_concurrent_stack_push_and_try_pop_lifo() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of String)()
        s.Push("First")
        s.Push("Second")

        Dim item1 As String = Nothing
        Dim item2 As String = Nothing
        Dim ok1 = s.TryPop(item1)
        Dim ok2 = s.TryPop(item2)

        Console.WriteLine(ok1 & "|" & item1 & "|" & ok2 & "|" & item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|Second|True|First"]);
}

#[test]
fn test_vb_concurrent_stack_push_range_array() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        s.PushRange(New Integer() {10, 20, 30})

        Dim item As Integer
        s.TryPop(item)
        Console.WriteLine(item & "|Count=" & s.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30|Count=2"]);
}

#[test]
fn test_vb_concurrent_stack_try_pop_range_array() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of String)()
        s.Push("A")
        s.Push("B")
        s.Push("C")

        Dim buffer(2) As String
        ' TryPopRange pops up to count elements into buffer array
        Dim poppedCount = s.TryPopRange(buffer, 0, 2)
        Console.WriteLine(poppedCount & "|" & buffer(0) & "|" & buffer(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|C|B"]);
}

#[test]
fn test_vb_concurrent_stack_try_peek() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Double)()
        s.Push(99.9)

        Dim peekVal As Double
        Dim ok = s.TryPeek(peekVal)
        Console.WriteLine(ok & "|" & peekVal & "|Count=" & s.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|99.9|Count=1"]);
}

#[test]
fn test_vb_concurrent_stack_try_pop_empty_returns_false() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        Dim item As Integer = 999
        Dim ok = s.TryPop(item)
        Console.WriteLine(ok & "|" & item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|0"]);
}

#[test]
fn test_vb_concurrent_stack_clear() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of String)()
        s.Push("X")
        s.Push("Y")
        s.Clear()
        Console.WriteLine(s.IsEmpty & "|" & s.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|0"]);
}

#[test]
fn test_vb_concurrent_stack_constructor_enumerable() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 3}
        Dim s As New ConcurrentStack(Of Integer)(list)
        Dim top As Integer
        s.TryPop(top)
        Console.WriteLine(top)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_concurrent_stack_to_array_snapshot() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        s.Push(10)
        s.Push(20)

        Dim arr As Integer() = s.ToArray()
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20,10"]);
}

#[test]
fn test_vb_concurrent_stack_multithreaded_push_pop() {
    let src = r#"
Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        Parallel.For(0, 100, Sub(i) s.Push(i))
        Console.WriteLine("Stack Count: " & s.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Stack Count: 100"]);
}

#[test]
fn test_vb_concurrent_stack_push_range_subslice() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of String)()
        Dim raw As String() = {"A", "B", "C", "D"}
        ' PushRange(items, offset, count)
        s.PushRange(raw, 1, 2)

        Dim top As String = Nothing
        s.TryPop(top)
        Console.WriteLine(top & "|Count=" & s.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["C|Count=1"]);
}

#[test]
fn test_vb_concurrent_stack_struct_elements() {
    let src = r#"
Imports System.Collections.Concurrent

Structure FrameInfo
    Public FrameId As Integer
    Public Symbol As String
End Structure

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of FrameInfo)()
        s.Push(New FrameInfo With {.FrameId = 1, .Symbol = "Main"})
        s.Push(New FrameInfo With {.FrameId = 2, .Symbol = "SubRoutine"})

        Dim info As FrameInfo
        s.TryPop(info)
        Console.WriteLine(info.FrameId & ":" & info.Symbol)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2:SubRoutine"]);
}

#[test]
fn test_vb_concurrent_stack_is_empty_property() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of String)()
        Console.WriteLine("Initial Empty: " & s.IsEmpty)
        s.Push("Data")
        Console.WriteLine("After Push Empty: " & s.IsEmpty)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Initial Empty: True", "After Push Empty: False"]
    );
}

#[test]
fn test_vb_concurrent_stack_producer_consumer_collection_interface() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim col As IProducerConsumerCollection(Of String) = New ConcurrentStack(Of String)()
        col.TryAdd("StackItem")
        Dim item As String = Nothing
        col.TryTake(item)
        Console.WriteLine(item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["StackItem"]);
}

#[test]
fn test_vb_concurrent_stack_null_elements() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of String)()
        s.Push(Nothing)

        Dim item As String = "NonNull"
        Dim ok = s.TryPop(item)
        Console.WriteLine(ok & "|" & (item Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_concurrent_stack_push_range_null_array_throws() {
    let src = r#"
Imports System
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        Try
            s.PushRange(Nothing)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null PushRange")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null PushRange"]
    );
}

#[test]
fn test_vb_concurrent_stack_try_pop_range_null_target_throws() {
    let src = r#"
Imports System
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        s.Push(1)
        Try
            s.TryPopRange(Nothing)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null Target PopRange")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null Target PopRange"]
    );
}

#[test]
fn test_vb_concurrent_stack_enumeration_top_to_bottom() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of String)()
        s.Push("Bottom")
        s.Push("Middle")
        s.Push("Top")

        Dim order = String.Join("->", s)
        Console.WriteLine(order)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Top->Middle->Bottom"]);
}

#[test]
fn test_vb_concurrent_stack_copy_to_array() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        s.Push(10)
        s.Push(20)

        Dim dest(3) As Integer
        s.CopyTo(dest, 1)
        Console.WriteLine(String.Join(",", dest))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0,20,10,0"]);
}

#[test]
fn test_vb_concurrent_stack_linq_queries() {
    let src = r#"
Imports System.Collections.Concurrent
Imports System.Linq

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        For i As Integer = 1 To 5 : s.Push(i) : Next
        Dim filtered = s.Where(Function(n) n > 2).ToList()
        Console.WriteLine(String.Join(",", filtered))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5,4,3"]);
}

#[test]
fn test_vb_concurrent_stack_try_peek_empty_returns_false() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of String)()
        Dim item As String = "Default"
        Dim ok = s.TryPeek(item)
        Console.WriteLine(ok & "|" & (item Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|True"]);
}
