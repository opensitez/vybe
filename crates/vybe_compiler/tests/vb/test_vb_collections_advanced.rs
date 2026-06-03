use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Collections — List, Dictionary, arrays, patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn list_basic_operations() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["3", "banana", "true"])
    );
}

#[test]
fn list_remove_and_indexof() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["1", "2", "30"]);
}

#[test]
fn list_foreach() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn dictionary_basic() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["10", "true", "false", "2"])
    );
}

#[test]
fn dictionary_remove() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, super::helpers::dotnet_expected_lines(&["1", "false"]));
}

#[test]
fn queue_operations() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["3", "first", "second", "2"]);
}

#[test]
fn stack_operations() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["3", "3", "2"]);
}

#[test]
fn array_declaration() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["30", "4"]);
}

#[test]
fn array_initializer() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn redim_array() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["4", "1", "5"]);
}

#[test]
fn list_clear() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["3", "0"]);
}

#[test]
fn object_in_list() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["3", "b"]);
}

#[test]
fn arraylist_basic_methods() {
    let out = run_vb(
        r#"
Imports System.Collections
Module M
    Sub Main()
        Dim al As New ArrayList()
        al.Add("A")
        al.Add("B")
        al.Add("C")
        Console.WriteLine(al.Count)
        Console.WriteLine(al.Item(0))
        Console.WriteLine(al.Item(2))
        Console.WriteLine(al.Contains("B"))
        Console.WriteLine(al.Contains("Z"))
        Console.WriteLine(al.IndexOf("B"))
        al.Remove("B")
        Console.WriteLine(al.Count)
        al.Insert(1, "X")
        Console.WriteLine(al.Item(1))
        al.Sort()
        Console.WriteLine(al.Item(0))
        al.Reverse()
        Console.WriteLine(al.Item(0))
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&[
            "3", "A", "C", "true", "false", "1", "2", "X", "A", "X"
        ])
    );
}

#[test]
fn arraylist_range_operations() {
    let out = run_vb(
        r#"
Imports System.Collections
Module M
    Sub Main()
        Dim al As New ArrayList()
        al.Add("A")
        al.Add("B")
        al.Add("C")
        al.Add("D")
        al.Add("E")

        ' InsertRange
        Dim ins As New ArrayList()
        ins.Add("X")
        ins.Add("Y")
        al.InsertRange(2, ins)
        Console.WriteLine(al.Count)
        Console.WriteLine(al.Item(2))
        Console.WriteLine(al.Item(3))

        ' RemoveRange
        al.RemoveRange(2, 2)
        Console.WriteLine(al.Count)
        Console.WriteLine(al.Item(2))

        ' GetRange
        Dim sub1 As ArrayList = al.GetRange(1, 3)
        Console.WriteLine(sub1.Count)
        Console.WriteLine(sub1.Item(0))

        ' SetRange
        Dim rep As New ArrayList()
        rep.Add("P")
        rep.Add("Q")
        al.SetRange(1, rep)
        Console.WriteLine(al.Item(1))
        Console.WriteLine(al.Item(2))

        ' Clone
        Dim cloned As ArrayList = al.Clone()
        Console.WriteLine(cloned.Count)
        cloned.Add("NEW")
        Console.WriteLine(cloned.Count)
        Console.WriteLine(al.Count)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["7", "X", "Y", "5", "C", "3", "B", "P", "Q", "5", "6", "5"]
    );
}

#[test]
fn arraylist_indexof_lastindexof() {
    let out = run_vb(
        r#"
Imports System.Collections
Module M
    Sub Main()
        Dim al As New ArrayList()
        al.Add("A")
        al.Add("B")
        al.Add("C")
        al.Add("D")
        al.Add("E")
        al.Add("B")
        Console.WriteLine(al.IndexOf("B"))
        Console.WriteLine(al.IndexOf("B", 2))
        Console.WriteLine(al.LastIndexOf("B"))
        Console.WriteLine(al.LastIndexOf("B", 3))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "5", "5", "1"]);
}

#[test]
fn arraylist_reverse_range() {
    let out = run_vb(
        r#"
Imports System.Collections
Module M
    Sub Main()
        Dim rv As New ArrayList()
        rv.Add(1)
        rv.Add(2)
        rv.Add(3)
        rv.Add(4)
        rv.Add(5)
        rv.Reverse(1, 3)
        Console.WriteLine(rv.Item(0))
        Console.WriteLine(rv.Item(1))
        Console.WriteLine(rv.Item(2))
        Console.WriteLine(rv.Item(3))
        Console.WriteLine(rv.Item(4))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "4", "3", "2", "5"]);
}

#[test]
fn arraylist_hof_methods() {
    let out = run_vb(
        r#"
Imports System.Collections
Module M
    Sub Main()
        Dim nums As New ArrayList()
        nums.Add(10)
        nums.Add(20)
        nums.Add(30)
        nums.Add(40)
        nums.Add(50)

        Dim fi As Integer = nums.FindIndex(Function(x) x > 25)
        Console.WriteLine(fi)

        Dim found As Object = nums.Find(Function(x) x > 25)
        Console.WriteLine(found)

        Console.WriteLine(nums.Exists(Function(x) x = 40))
        Console.WriteLine(nums.Exists(Function(x) x = 99))
        Console.WriteLine(nums.TrueForAll(Function(x) x > 0))
        Console.WriteLine(nums.TrueForAll(Function(x) x > 15))

        Dim doubled As ArrayList = nums.ConvertAll(Function(x) x * 2)
        Console.WriteLine(doubled.Count)
        Console.WriteLine(doubled.Item(0))

        Dim nums2 As New ArrayList()
        nums2.Add(1)
        nums2.Add(2)
        nums2.Add(3)
        nums2.Add(4)
        nums2.Add(5)
        Dim removed As Integer = nums2.RemoveAll(Function(x) x > 3)
        Console.WriteLine(removed)
        Console.WriteLine(nums2.Count)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&[
            "2", "30", "true", "false", "true", "false", "5", "20", "2", "3"
        ])
    );
}

#[test]
fn concurrent_collections_basic() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent
Sub Main()
    Dim d As New ConcurrentDictionary
    d.TryAdd("key1", "value1")
    d.TryAdd("key2", "value2")
    Console.WriteLine(d.ContainsKey("key1"))
    Console.WriteLine(d.ContainsKey("missing"))

    Dim q As New ConcurrentQueue
    q.Enqueue(100)
    q.Enqueue(200)
    Console.WriteLine(q.Count)

    Dim s As New ConcurrentStack
    s.Push(10)
    s.Push(20)
    Console.WriteLine(s.Count)
End Sub
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["true", "false", "2", "2"])
    );
}
