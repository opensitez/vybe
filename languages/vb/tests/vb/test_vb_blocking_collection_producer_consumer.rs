use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: BlockingCollection(Of T) & Producer-Consumer
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_blocking_collection_add_take() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of String)()
        bc.Add("Item1")
        bc.Add("Item2")
        Console.WriteLine(bc.Take())
        Console.WriteLine(bc.Take())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Item1", "Item2"]);
}

#[test]
fn test_vb_blocking_collection_try_add_try_take() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)(boundedCapacity:=2)
        Console.WriteLine(bc.TryAdd(1))
        Console.WriteLine(bc.TryAdd(2))
        Console.WriteLine(bc.TryAdd(3)) ' Exceeds capacity

        Dim item As Integer
        Console.WriteLine(bc.TryTake(item))
        Console.WriteLine(item)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True", "False", "True", "1"]);
}

#[test]
fn test_vb_blocking_collection_complete_adding_consuming_enumerable() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of Integer)()
        bc.Add(10)
        bc.Add(20)
        bc.CompleteAdding()

        Console.WriteLine(bc.IsAddingCompleted)
        Dim sum As Integer = 0
        For Each val In bc.GetConsumingEnumerable()
            sum += val
        Next
        Console.WriteLine(sum)
        Console.WriteLine(bc.IsCompleted)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "30", "True"]);
}

#[test]
fn test_vb_blocking_collection_with_concurrent_stack() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bc As New BlockingCollection(Of String)(New ConcurrentStack(Of String)())
        bc.Add("First")
        bc.Add("Second")
        Console.WriteLine(bc.Take())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Second"]);
}
