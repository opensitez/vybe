use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ConcurrentBag(Of T) Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_concurrent_bag_add_try_take() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bag As New ConcurrentBag(Of Integer)()
        bag.Add(10)
        bag.Add(20)
        Dim item As Integer
        Dim ok As Boolean = bag.TryTake(item)
        Console.WriteLine(ok)
        Console.WriteLine(bag.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "1"]);
}

#[test]
fn test_vb_concurrent_bag_try_peek() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bag As New ConcurrentBag(Of String)()
        bag.Add("Item")
        Dim peeked As String = Nothing
        Dim ok As Boolean = bag.TryPeek(peeked)
        Console.WriteLine(ok)
        Console.WriteLine(peeked)
        Console.WriteLine(bag.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "Item", "1"]);
}

#[test]
fn test_vb_concurrent_bag_is_empty_property() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bag As New ConcurrentBag(Of Integer)()
        Console.WriteLine(bag.IsEmpty)
        bag.Add(1)
        Console.WriteLine(bag.IsEmpty)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_concurrent_bag_to_array_snapshot() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bag As New ConcurrentBag(Of Integer)()
        bag.Add(1)
        bag.Add(2)
        Dim arr As Integer() = bag.ToArray()
        Console.WriteLine(arr.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_concurrent_bag_clear() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim bag As New ConcurrentBag(Of Integer)()
        bag.Add(1)
        bag.Add(2)
        bag.Clear()
        Console.WriteLine(bag.IsEmpty)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
