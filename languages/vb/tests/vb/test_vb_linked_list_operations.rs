use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LinkedList(Of T) & LinkedListNode(Of T) Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linked_list_add_first_add_last() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ll As New LinkedList(Of String)()
        ll.AddLast("Middle")
        ll.AddFirst("First")
        ll.AddLast("Last")
        Console.WriteLine(String.Join(",", ll))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First,Middle,Last"]);
}

#[test]
fn test_vb_linked_list_node_navigation_next_previous() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ll As New LinkedList(Of Integer)()
        ll.AddLast(10)
        ll.AddLast(20)
        ll.AddLast(30)

        Dim node As LinkedListNode(Of Integer) = ll.First.Next
        Console.WriteLine(node.Value)
        Console.WriteLine(node.Previous.Value)
        Console.WriteLine(node.Next.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20", "10", "30"]);
}

#[test]
fn test_vb_linked_list_add_before_add_after_node() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ll As New LinkedList(Of String)()
        Dim node As LinkedListNode(Of String) = ll.AddLast("Target")
        ll.AddBefore(node, "Before")
        ll.AddAfter(node, "After")
        Console.WriteLine(String.Join(",", ll))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Before,Target,After"]);
}

#[test]
fn test_vb_linked_list_find_and_find_last() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ll As New LinkedList(Of Integer)()
        ll.AddLast(10)
        ll.AddLast(20)
        ll.AddLast(10)

        Dim first10 As LinkedListNode(Of Integer) = ll.Find(10)
        Dim last10 As LinkedListNode(Of Integer) = ll.FindLast(10)
        Console.WriteLine(Object.ReferenceEquals(first10, ll.First))
        Console.WriteLine(Object.ReferenceEquals(last10, ll.Last))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_linked_list_remove_first_last() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ll As New LinkedList(Of Integer)()
        ll.AddLast(1)
        ll.AddLast(2)
        ll.AddLast(3)
        ll.RemoveFirst()
        ll.RemoveLast()
        Console.WriteLine(ll.Count)
        Console.WriteLine(ll.First.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1", "2"]);
}

#[test]
fn test_vb_linked_list_remove_node() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ll As New LinkedList(Of String)()
        Dim n1 As LinkedListNode(Of String) = ll.AddLast("A")
        Dim n2 As LinkedListNode(Of String) = ll.AddLast("B")
        ll.Remove(n1)
        Console.WriteLine(ll.First.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B"]);
}

#[test]
fn test_vb_linked_list_contains_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ll As New LinkedList(Of Integer)()
        ll.AddLast(100)
        Console.WriteLine(ll.Contains(100))
        Console.WriteLine(ll.Contains(200))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_linked_list_copy_to_array() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ll As New LinkedList(Of Integer)()
        ll.AddLast(10)
        ll.AddLast(20)
        Dim arr(1) As Integer
        ll.CopyTo(arr, 0)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_linked_list_clear() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ll As New LinkedList(Of String)()
        ll.AddLast("Item")
        ll.Clear()
        Console.WriteLine(ll.Count)
        Console.WriteLine(ll.First Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0", "True"]);
}

#[test]
fn test_vb_linked_list_node_list_property() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim ll As New LinkedList(Of String)()
        Dim node As LinkedListNode(Of String) = ll.AddLast("Test")
        Console.WriteLine(Object.ReferenceEquals(node.List, ll))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
