use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Collections — List, Dictionary, arrays, patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn list_basic_operations() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim items As New List(Of String)
        items.Add("apple")
        items.Add("banana")
        items.Add("cherry")
        Console.WriteLine(items.Count)
        Console.WriteLine(items.Item(1))
        Console.WriteLine(items.Contains("banana"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3", "banana", "true"]);
}

#[test]
fn list_remove_and_indexof() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim nums As New List(Of Integer)
        nums.Add(10)
        nums.Add(20)
        nums.Add(30)
        Console.WriteLine(nums.IndexOf(20))
        nums.Remove(20)
        Console.WriteLine(nums.Count)
        Console.WriteLine(nums.Item(1))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["1", "2", "30"]);
}

#[test]
fn list_foreach() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim items As New List(Of String)
        items.Add("a")
        items.Add("b")
        items.Add("c")
        Dim result As String = ""
        For Each item As String In items
            result = result & item
        Next
        Console.WriteLine(result)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn dictionary_basic() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim d As New Dictionary(Of String, Integer)
        d.Add("x", 10)
        d.Add("y", 20)
        Console.WriteLine(d.Item("x"))
        Console.WriteLine(d.ContainsKey("y"))
        Console.WriteLine(d.ContainsKey("z"))
        Console.WriteLine(d.Count)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["10", "true", "false", "2"]);
}

#[test]
fn dictionary_remove() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim d As New Dictionary(Of String, Integer)
        d.Add("a", 1)
        d.Add("b", 2)
        d.Remove("a")
        Console.WriteLine(d.Count)
        Console.WriteLine(d.ContainsKey("a"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["1", "false"]);
}

#[test]
fn queue_operations() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim q As New Queue(Of String)
        q.Enqueue("first")
        q.Enqueue("second")
        q.Enqueue("third")
        Console.WriteLine(q.Count)
        Console.WriteLine(q.Dequeue())
        Console.WriteLine(q.Peek())
        Console.WriteLine(q.Count)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3", "first", "second", "2"]);
}

#[test]
fn stack_operations() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim s As New Stack(Of Integer)
        s.Push(1)
        s.Push(2)
        s.Push(3)
        Console.WriteLine(s.Count)
        Console.WriteLine(s.Pop())
        Console.WriteLine(s.Peek())
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3", "3", "2"]);
}

#[test]
fn array_declaration() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim arr(4) As Integer
        arr(0) = 10
        arr(1) = 20
        arr(2) = 30
        arr(3) = 40
        arr(4) = 50
        Console.WriteLine(arr(2))
        Console.WriteLine(UBound(arr))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["30", "4"]);
}

#[test]
fn array_initializer() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim arr() As Integer = {5, 10, 15, 20}
        Dim sum As Integer = 0
        For Each n As Integer In arr
            sum = sum + n
        Next
        Console.WriteLine(sum)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn redim_array() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim arr(2) As Integer
        arr(0) = 1
        arr(1) = 2
        arr(2) = 3
        ReDim Preserve arr(4)
        arr(3) = 4
        arr(4) = 5
        Console.WriteLine(UBound(arr))
        Console.WriteLine(arr(0))
        Console.WriteLine(arr(4))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["4", "1", "5"]);
}

#[test]
fn list_clear() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim items As New List(Of Integer)
        items.Add(1)
        items.Add(2)
        items.Add(3)
        Console.WriteLine(items.Count)
        items.Clear()
        Console.WriteLine(items.Count)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3", "0"]);
}

#[test]
fn object_in_list() {
    let out = run_vb(r#"
Class Item
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
End Class

Module M
    Sub Main()
        Dim items As New List(Of Item)
        items.Add(New Item("a"))
        items.Add(New Item("b"))
        items.Add(New Item("c"))
        Console.WriteLine(items.Count)
        Console.WriteLine(items.Item(1).Name)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3", "b"]);
}
