use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Collections.Concurrent.ConcurrentQueue FIFO
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_concurrent_queue_enqueue_and_try_dequeue_fifo() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of String)()
        q.Enqueue("First")
        q.Enqueue("Second")

        Dim item1 As String = Nothing
        Dim item2 As String = Nothing
        Dim ok1 = q.TryDequeue(item1)
        Dim ok2 = q.TryDequeue(item2)

        Console.WriteLine(ok1 & "|" & item1 & "|" & ok2 & "|" & item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|First|True|Second"]);
}

#[test]
fn test_vb_concurrent_queue_try_peek() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        q.Enqueue(42)

        Dim peekVal As Integer
        Dim ok = q.TryPeek(peekVal)
        Console.WriteLine(ok & "|" & peekVal & "|Count=" & q.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|42|Count=1"]);
}

#[test]
fn test_vb_concurrent_queue_try_dequeue_empty_returns_false() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of String)()
        Dim item As String = Nothing
        Dim ok = q.TryDequeue(item)
        Console.WriteLine(ok & "|" & (item Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|True"]);
}

#[test]
fn test_vb_concurrent_queue_multithreaded_producer_consumer() {
    let src = r#"
Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        Parallel.For(0, 50, Sub(i) q.Enqueue(i))

        Dim count = 0
        Dim val As Integer
        While q.TryDequeue(val)
            count += 1
        End While

        Console.WriteLine("Dequeued Total: " & count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Dequeued Total: 50"]);
}

#[test]
fn test_vb_concurrent_queue_clear() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        q.Enqueue(1)
        q.Enqueue(2)
        q.Clear()
        Console.WriteLine(q.IsEmpty & "|" & q.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|0"]);
}

#[test]
fn test_vb_concurrent_queue_constructor_collection() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim list As New List(Of String) From {"A", "B", "C"}
        Dim q As New ConcurrentQueue(Of String)(list)
        Dim item As String = Nothing
        q.TryDequeue(item)
        Console.WriteLine(q.Count & "|" & item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|A"]);
}

#[test]
fn test_vb_concurrent_queue_to_array_snapshot() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        q.Enqueue(100)
        q.Enqueue(200)

        Dim snapshot As Integer() = q.ToArray()
        Console.WriteLine(String.Join(",", snapshot))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100,200"]);
}

#[test]
fn test_vb_concurrent_queue_copy_to_array() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of String)()
        q.Enqueue("X")
        q.Enqueue("Y")

        Dim target(3) As String
        q.CopyTo(target, 1)
        Console.WriteLine(String.Join(",", target))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec![",X,Y,"]);
}

#[test]
fn test_vb_concurrent_queue_struct_elements() {
    let src = r#"
Imports System.Collections.Concurrent

Structure TaskRecord
    Public Id As Integer
    Public TaskName As String
End Structure

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of TaskRecord)()
        q.Enqueue(New TaskRecord With {.Id = 1, .TaskName = "Job1"})

        Dim rec As TaskRecord
        q.TryDequeue(rec)
        Console.WriteLine(rec.Id & ":" & rec.TaskName)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:Job1"]);
}

#[test]
fn test_vb_concurrent_queue_enumeration_snapshot_semantic() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        q.Enqueue(1)
        q.Enqueue(2)

        Dim res = ""
        For Each item In q
            res &= item & ","
            If item = 1 Then q.Enqueue(3) ' Mutation does not affect active enumerator snapshot!
        Next
        Console.WriteLine(res.TrimEnd(","c) & "|Count=" & q.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2|Count=3"]);
}

#[test]
fn test_vb_concurrent_queue_null_element_allowed() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of String)()
        q.Enqueue(Nothing)

        Dim item As String = "NonNull"
        Dim ok = q.TryDequeue(item)
        Console.WriteLine(ok & "|" & (item Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_concurrent_queue_is_empty_property() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        Console.WriteLine("Initially Empty: " & q.IsEmpty)
        q.Enqueue(10)
        Console.WriteLine("After Enqueue Empty: " & q.IsEmpty)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Initially Empty: True", "After Enqueue Empty: False"]
    );
}

#[test]
fn test_vb_concurrent_queue_iconcurrentcollection_implementation() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim col As IProducerConsumerCollection(Of String) = New ConcurrentQueue(Of String)()
        col.TryAdd("ItemA")
        Dim item As String = Nothing
        col.TryTake(item)
        Console.WriteLine(item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ItemA"]);
}

#[test]
fn test_vb_concurrent_queue_interleaved_enqueue_dequeue() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        q.Enqueue(1)
        q.Enqueue(2)
        Dim v1 As Integer
        q.TryDequeue(v1)

        q.Enqueue(3)
        Dim v2 As Integer, v3 As Integer
        q.TryDequeue(v2)
        q.TryDequeue(v3)

        Console.WriteLine(v1 & "|" & v2 & "|" & v3)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|2|3"]);
}

#[test]
fn test_vb_concurrent_queue_try_peek_empty_returns_false() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        Dim peekVal As Integer = 999
        Dim ok = q.TryPeek(peekVal)
        Console.WriteLine(ok & "|" & peekVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|0"]);
}

#[test]
fn test_vb_concurrent_queue_generic_object_payload() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Object)()
        q.Enqueue(100)
        q.Enqueue("StringPayload")

        Dim o1 As Object = Nothing, o2 As Object = Nothing
        q.TryDequeue(o1)
        q.TryDequeue(o2)

        Console.WriteLine(o1.GetType().Name & "|" & o2.GetType().Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32|String"]);
}

#[test]
fn test_vb_concurrent_queue_linq_query_filtering() {
    let src = r#"
Imports System.Collections.Concurrent
Imports System.Linq

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        For i As Integer = 1 To 10 : q.Enqueue(i) : Next
        Dim evens = q.Where(Function(n) n Mod 2 = 0).ToList()
        Console.WriteLine(String.Join(",", evens))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2,4,6,8,10"]);
}

#[test]
fn test_vb_concurrent_queue_multiple_dequeues_in_parallel() {
    let src = r#"
Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        For i As Integer = 1 To 100 : q.Enqueue(i) : Next

        Dim sum = 0
        Dim lockObj As New Object()
        Parallel.For(0, 100, Sub(i)
            Dim item As Integer
            If q.TryDequeue(item) Then
                SyncLock lockObj
                    sum += item
                End SyncLock
            End If
        End Sub)
        Console.WriteLine(sum & "|QueueEmpty=" & q.IsEmpty)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5050|QueueEmpty=True"]);
}

#[test]
fn test_vb_concurrent_queue_copy_to_null_target_throws() {
    let src = r#"
Imports System
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        q.Enqueue(1)
        Try
            q.CopyTo(Nothing, 0)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null Target CopyTo")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null Target CopyTo"]
    );
}

#[test]
fn test_vb_concurrent_queue_copy_to_invalid_index_throws() {
    let src = r#"
Imports System
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        q.Enqueue(1)
        Dim target(2) As Integer
        Try
            q.CopyTo(target, -1)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ArgumentOutOfRangeException Caught on CopyTo Invalid Index")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentOutOfRangeException Caught on CopyTo Invalid Index"]
    );
}
