use super::helpers::run_vb;

#[test]
fn collection_api_matrix_list_add_insert_remove_capacity() {
    let out = run_vb(
        r#"
Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 4}
        list.Insert(2, 3)
        list.RemoveAt(0)

        Console.WriteLine(list.Count)
        Console.WriteLine(list(0))
        Console.WriteLine(list(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "2", "4"]);
}

#[test]
fn collection_api_matrix_list_find_and_exists() {
    let out = run_vb(
        r#"
Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {10, 20, 30, 40}
        Dim has30 As Boolean = values.Contains(30)
        Dim indexOf10 As Integer = values.IndexOf(10)
        Dim indexOf50 As Integer = values.IndexOf(50)

        Console.WriteLine(has30)
        Console.WriteLine(indexOf10)
        Console.WriteLine(indexOf50)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "0", "-1"]);
}

#[test]
fn collection_api_matrix_dictionary_lookup_contract() {
    let out = run_vb(
        r#"
Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim m As New Dictionary(Of String, Integer)()
        m.Add("a", 1)
        m("b") = 2

        Dim value As Integer = -1
        Dim hasC As Boolean = m.TryGetValue("c", value)
        Dim hasA As Boolean = m.TryGetValue("a", value)

        Console.WriteLine(m.ContainsKey("a"))
        Console.WriteLine(hasC)
        Console.WriteLine(hasA)
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False", "True", "1"]);
}

#[test]
fn collection_api_matrix_dictionary_enumeration_order_stable() {
    let out = run_vb(
        r#"
Imports System
Imports System.Collections.Generic
Imports System.Text

Module M
    Sub Main()
        Dim map As New Dictionary(Of Integer, String)()
        map.Add(2, "b")
        map.Add(1, "a")
        map.Add(3, "c")

        Dim sb As New StringBuilder()
        For Each pair In map
            sb.Append(pair.Key).Append(":").Append(pair.Value).Append(",")
        Next

        Console.WriteLine(sb.ToString().Contains("1:a"))
        Console.WriteLine(sb.ToString().Contains("3:c"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn collection_api_matrix_hashset_uniqueness_semantics() {
    let out = run_vb(
        r#"
Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim set As New HashSet(Of Integer)()

        set.Add(1)
        set.Add(1)
        set.Add(2)

        Console.WriteLine(set.Count)
        Console.WriteLine(set.Contains(2))
        Console.WriteLine(set.Remove(1))
        Console.WriteLine(set.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "True", "True", "1"]);
}

#[test]
fn collection_api_matrix_set_intersection_like() {
    let out = run_vb(
        r#"
Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim left As New HashSet(Of Integer)({1, 2, 3, 4})
        Dim right As New HashSet(Of Integer)({3, 4, 5})
        left.IntersectWith(right)

        Dim ordered As New List(Of Integer)(left)
        ordered.Sort()

        Console.WriteLine(String.Join(",", ordered))
        Console.WriteLine(left.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3,4", "2"]);
}

#[test]
fn collection_api_matrix_queue_fifo_behavior() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim queue As New Queue(Of String)()

        queue.Enqueue("a")
        queue.Enqueue("b")
        queue.Enqueue("c")

        Dim a As String = queue.Dequeue()
        Dim b As String = queue.Dequeue()

        Console.WriteLine(a)
        Console.WriteLine(b)
        Console.WriteLine(queue.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["a", "b", "1"]);
}

#[test]
fn collection_api_matrix_stack_lifo_behavior() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim stack As New Stack(Of Integer)()
        stack.Push(1)
        stack.Push(2)
        stack.Push(3)

        Console.WriteLine(stack.Pop())
        Console.WriteLine(stack.Peek())
        Console.WriteLine(stack.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "2", "2"]);
}

#[test]
fn collection_api_matrix_array_to_list_and_back() {
    let out = run_vb(
        r#"
Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim source As Integer() = {1, 2, 3}
        Dim list As New List(Of Integer)(source)
        list.Add(4)
        Dim arr As Integer() = list.ToArray()

        Console.WriteLine(list.Count)
        Console.WriteLine(arr.Length)
        Console.WriteLine(arr(3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["4", "4", "4"]);
}
