use super::helpers::run_vb;

#[test]
fn concurrent_dictionary_add_and_retrieve() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim values As New ConcurrentDictionary(Of String, Integer)()

        Dim added As Boolean = values.TryAdd("alpha", 7)
        Dim value As Integer = 0
        Dim retrieved As Boolean = values.TryGetValue("alpha", value)

        Console.WriteLine(added)
        Console.WriteLine(retrieved)
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "7"]);
}

#[test]
fn concurrent_dictionary_rejects_duplicate_key() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim values As New ConcurrentDictionary(Of String, Integer)()

        values.TryAdd("alpha", 1)
        Dim duplicate As Boolean = values.TryAdd("alpha", 2)

        Console.WriteLine(duplicate)
        Console.WriteLine(values("alpha"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "1"]);
}

#[test]
fn concurrent_dictionary_get_or_add_factory() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim values As New ConcurrentDictionary(Of String, Integer)()

        Dim first As Integer = values.GetOrAdd("k", Function(key As String) 100)
        Dim second As Integer = values.GetOrAdd("k", Function(key As String) 200)

        Console.WriteLine(first)
        Console.WriteLine(second)
        Console.WriteLine(values.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["100", "100", "1"]);
}

#[test]
fn concurrent_dictionary_add_or_update() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim values As New ConcurrentDictionary(Of String, Integer)()

        Dim updated As Integer = values.AddOrUpdate("score", 3, Function(key As String, oldValue As Integer) oldValue + 2)
        Console.WriteLine(updated)
        Console.WriteLine(values("score"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn concurrent_dictionary_add_or_update_existing_key() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim values As New ConcurrentDictionary(Of String, Integer)()

        values.TryAdd("score", 10)
        Dim updated As Integer = values.AddOrUpdate("score", 3, Function(key As String, oldValue As Integer) oldValue + 1)

        Console.WriteLine(updated)
        Console.WriteLine(values("score"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["11", "11"]);
}

#[test]
fn concurrent_dictionary_try_update() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim values As New ConcurrentDictionary(Of String, Integer)()

        values.TryAdd("state", 1)
        Dim changed As Boolean = values.TryUpdate("state", 4, 1)
        Dim failed As Boolean = values.TryUpdate("state", 9, 1)

        Console.WriteLine(changed)
        Console.WriteLine(failed)
        Console.WriteLine(values("state"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False", "4"]);
}

#[test]
fn concurrent_dictionary_remove_value() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim values As New ConcurrentDictionary(Of String, Integer)()

        values.TryAdd("temp", 5)
        Dim removed As Integer = 0
        Dim success As Boolean = values.TryRemove("temp", removed)

        Console.WriteLine(success)
        Console.WriteLine(removed)
        Console.WriteLine(values.ContainsKey("temp"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "5", "False"]);
}

#[test]
fn concurrent_queue_fifo_order() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        q.Enqueue(1)
        q.Enqueue(2)
        q.Enqueue(3)

        Dim first As Integer = 0
        Dim second As Integer = 0
        Dim third As Integer = 0

        Console.WriteLine(q.TryPeek(first))
        Console.WriteLine(first)

        Console.WriteLine(q.TryDequeue(first))
        Console.WriteLine(first)
        Console.WriteLine(q.TryDequeue(second))
        Console.WriteLine(second)
        Console.WriteLine(q.TryDequeue(third))
        Console.WriteLine(third)
        Console.WriteLine(q.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(
        out,
        vec!["True", "1", "True", "1", "True", "2", "True", "3", "0"]
    );
}

#[test]
fn concurrent_queue_dequeue_empty_queue_returns_false() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        Dim value As Integer = 42

        Console.WriteLine(q.TryDequeue(value))
        Console.WriteLine(value)
        Console.WriteLine(q.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "42", "0"]);
}

#[test]
fn concurrent_stack_lifo_behavior() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        s.Push(1)
        s.Push(2)

        Dim top As Integer = 0
        Console.WriteLine(s.TryPeek(top))
        Console.WriteLine(top)

        Dim removed As Integer = 0
        Console.WriteLine(s.TryPop(removed))
        Console.WriteLine(removed)
        Console.WriteLine(s.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "2", "True", "2", "1"]);
}

#[test]
fn concurrent_bag_add_take_and_count() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim bag As New ConcurrentBag(Of String)()
        bag.Add("left")
        bag.Add("right")

        Dim countBefore As Integer = bag.Count
        Dim head As String = ""

        Console.WriteLine(countBefore)
        Console.WriteLine(bag.TryPeek(head))
        Console.WriteLine(head = "left" OrElse head = "right")

        Dim taken As String = ""
        Console.WriteLine(bag.TryTake(taken))
        Console.WriteLine(bag.Count)
        Console.WriteLine(taken = "left" OrElse taken = "right")
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "True", "True", "True", "1", "True"]);
}
