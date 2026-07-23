use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Collections.Concurrent.BlockingCollection Producer-Consumer
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_blocking_collection_add_and_take() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of String)()
        bc.Add("Item1")
        Dim item = bc.Take()
        Console.WriteLine(item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Item1"]);
}

#[test]
fn test_vb_blocking_collection_try_take_with_timeout() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()
        Dim item As Integer
        Dim ok = bc.TryTake(item, millisecondsTimeout:=10)
        Console.WriteLine(ok & "|" & item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|0"]);
}

#[test]
fn test_vb_blocking_collection_complete_adding_enumeration() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()
        bc.Add(10)
        bc.Add(20)
        bc.CompleteAdding()

        Dim list As New System.Collections.Generic.List(Of Integer)()
        For Each item In bc.GetConsumingEnumerable()
            list.Add(item)
        Next
        Console.WriteLine(String.Join(",", list) & "|IsCompleted=" & bc.IsCompleted)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20|IsCompleted=True"]);
}

#[test]
fn test_vb_blocking_collection_add_after_complete_adding_throws() {
    let src = r#"
Imports System
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of String)()
        bc.CompleteAdding()
        Try
            bc.Add("AfterComplete")
        Catch ex As InvalidOperationException
            Console.WriteLine("InvalidOperationException Caught on Add After Complete")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidOperationException Caught on Add After Complete"]
    );
}

#[test]
fn test_vb_blocking_collection_bounded_capacity() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)(boundedCapacity:=2)
        Dim added1 = bc.TryAdd(1)
        Dim added2 = bc.TryAdd(2)
        Dim added3 = bc.TryAdd(3, millisecondsTimeout:=10) ' Exceeds bounded capacity!

        Console.WriteLine(added1 & "|" & added2 & "|" & added3)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|False"]);
}

#[test]
fn test_vb_blocking_collection_wrapping_concurrent_stack_lifo() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim stack As New ConcurrentStack(Of String)()
        Dim bc As New BlockingCollection(Of String)(stack)
        bc.Add("First")
        bc.Add("Second")

        ' Takes from underlying ConcurrentStack in LIFO order!
        Dim item1 = bc.Take()
        Dim item2 = bc.Take()
        Console.WriteLine(item1 & "|" & item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Second|First"]);
}

#[test]
fn test_vb_blocking_collection_take_from_any_multiple_collections() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc1 As New BlockingCollection(Of Integer)()
        Dim bc2 As New BlockingCollection(Of Integer)()
        bc2.Add(500)

        Dim item As Integer
        Dim idx = BlockingCollection(Of Integer).TakeFromAny(New BlockingCollection(Of Integer)() {bc1, bc2}, item)
        Console.WriteLine("Index: " & idx & "|Item: " & item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Index: 1|Item: 500"]);
}

#[test]
fn test_vb_blocking_collection_add_to_any_bounded() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc1 As New BlockingCollection(Of Integer)(boundedCapacity:=1)
        Dim bc2 As New BlockingCollection(Of Integer)(boundedCapacity:=1)
        bc1.Add(10)

        ' Adding to collection array routes to bc2 since bc1 is full!
        Dim idx = BlockingCollection(Of Integer).AddToAny(New BlockingCollection(Of Integer)() {bc1, bc2}, 20)
        Console.WriteLine("Added to Collection Index: " & idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Added to Collection Index: 1"]);
}

#[test]
fn test_vb_blocking_collection_is_adding_completed_property() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()
        Console.WriteLine("Before Complete: " & bc.IsAddingCompleted)
        bc.CompleteAdding()
        Console.WriteLine("After Complete: " & bc.IsAddingCompleted)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Before Complete: False", "After Complete: True"]
    );
}

#[test]
fn test_vb_blocking_collection_take_on_empty_completed_throws() {
    let src = r#"
Imports System
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of String)()
        bc.CompleteAdding()
        Try
            bc.Take()
        Catch ex As InvalidOperationException
            Console.WriteLine("InvalidOperationException Caught on Take Empty Completed")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidOperationException Caught on Take Empty Completed"]
    );
}

#[test]
fn test_vb_blocking_collection_dispose_disposes_underlying() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()
        bc.Dispose()
        Console.WriteLine("BlockingCollection Disposed")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["BlockingCollection Disposed"]);
}

#[test]
fn test_vb_blocking_collection_multithreaded_producer_consumer_pipeline() {
    let src = r#"
Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()

        Dim producer = Task.Run(Sub()
            For i As Integer = 1 To 5 : bc.Add(i) : Next
            bc.CompleteAdding()
        End Sub)

        Dim consumerSum = 0
        Dim consumer = Task.Run(Sub()
            For Each item In bc.GetConsumingEnumerable()
                consumerSum += item
            Next
        End Sub)

        Task.WaitAll(producer, consumer)
        Console.WriteLine("Consumer Sum: " & consumerSum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Consumer Sum: 15"]);
}

#[test]
fn test_vb_blocking_collection_to_array_snapshot() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of String)()
        bc.Add("Alpha")
        bc.Add("Beta")
        Dim snapshot As String() = bc.ToArray()
        Console.WriteLine(String.Join(",", snapshot) & "|Count=" & bc.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alpha,Beta|Count=2"]);
}

#[test]
fn test_vb_blocking_collection_struct_elements() {
    let src = r#"
Imports System.Collections.Concurrent

Structure LogEntry
    Public Level As Integer
    Public Message As String
End Structure

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of LogEntry)()
        bc.Add(New LogEntry With {.Level = 1, .Message = "InfoMsg"})

        Dim entry = bc.Take()
        Console.WriteLine(entry.Level & ":" & entry.Message)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:InfoMsg"]);
}

#[test]
fn test_vb_blocking_collection_bounded_capacity_zero_throws() {
    let src = r#"
Imports System
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Try
            Dim bc As New BlockingCollection(Of Integer)(boundedCapacity:=0)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ArgumentOutOfRangeException Caught on 0 Capacity")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentOutOfRangeException Caught on 0 Capacity"]
    );
}

#[test]
fn test_vb_blocking_collection_try_add_with_cancellation_token() {
    let src = r#"
Imports System.Collections.Concurrent
Imports System.Threading

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)(boundedCapacity:=1)
        bc.Add(100)

        Dim cts As New CancellationTokenSource()
        cts.Cancel()

        Try
            bc.Add(200, cts.Token)
        Catch ex As OperationCanceledException
            Console.WriteLine("OperationCanceledException Caught on Add")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["OperationCanceledException Caught on Add"]
    );
}

#[test]
fn test_vb_blocking_collection_try_take_with_cancellation_token() {
    let src = r#"
Imports System.Collections.Concurrent
Imports System.Threading

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()
        Dim cts As New CancellationTokenSource()
        cts.Cancel()

        Try
            Dim item = bc.Take(cts.Token)
        Catch ex As OperationCanceledException
            Console.WriteLine("OperationCanceledException Caught on Take")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["OperationCanceledException Caught on Take"]
    );
}

#[test]
fn test_vb_blocking_collection_wrapping_concurrent_bag() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bag As New ConcurrentBag(Of String)()
        Dim bc As New BlockingCollection(Of String)(bag)
        bc.Add("BagItem1")
        bc.Add("BagItem2")

        Dim count = 0
        While bc.Count > 0
            Dim item = bc.Take()
            count += 1
        End While
        Console.WriteLine("Bag Collection Cleared Count: " & count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Bag Collection Cleared Count: 2"]);
}

#[test]
fn test_vb_blocking_collection_get_consuming_enumerable_multiple_times() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()
        bc.Add(1)
        bc.Add(2)

        Dim enumr = bc.GetConsumingEnumerable().GetEnumerator()
        enumr.MoveNext()
        Console.WriteLine("Consumed First: " & enumr.Current & "|Remaining Count=" & bc.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Consumed First: 1|Remaining Count=1"]);
}

#[test]
fn test_vb_blocking_collection_take_from_any_timeout_returns_minus_one() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc1 As New BlockingCollection(Of Integer)()
        Dim bc2 As New BlockingCollection(Of Integer)()

        Dim item As Integer
        ' Timeout after 10ms when both collections are empty
        Dim idx = BlockingCollection(Of Integer).TryTakeFromAny(New BlockingCollection(Of Integer)() {bc1, bc2}, item, millisecondsTimeout:=10)
        Console.WriteLine(idx & "|" & item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1|0"]);
}
