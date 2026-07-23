use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ConcurrentDictionary(Of K, V) Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_concurrent_dict_try_add_try_get_value() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cd As New ConcurrentDictionary(Of String, Integer)()
        Dim ok1 As Boolean = cd.TryAdd("Key1", 100)
        Dim ok2 As Boolean = cd.TryAdd("Key1", 200)
        Console.WriteLine(ok1)
        Console.WriteLine(ok2)

        Dim val As Integer
        cd.TryGetValue("Key1", val)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False", "100"]);
}

#[test]
fn test_vb_concurrent_dict_get_or_add_value_factory() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cd As New ConcurrentDictionary(Of Integer, String)()
        Dim val1 As String = cd.GetOrAdd(1, Function(k) "Value_" & k)
        Dim val2 As String = cd.GetOrAdd(1, Function(k) "NewValue_" & k)
        Console.WriteLine(val1)
        Console.WriteLine(val2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Value_1", "Value_1"]);
}

#[test]
fn test_vb_concurrent_dict_add_or_update_factory() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cd As New ConcurrentDictionary(Of String, Integer)()
        Dim v1 As Integer = cd.AddOrUpdate("Counter", 1, Function(k, oldV) oldV + 1)
        Dim v2 As Integer = cd.AddOrUpdate("Counter", 1, Function(k, oldV) oldV + 1)
        Console.WriteLine(v1)
        Console.WriteLine(v2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1", "2"]);
}

#[test]
fn test_vb_concurrent_dict_try_update() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cd As New ConcurrentDictionary(Of String, Integer)()
        cd.TryAdd("K", 10)
        Dim okWrong As Boolean = cd.TryUpdate("K", 99, 5)
        Dim okRight As Boolean = cd.TryUpdate("K", 99, 10)
        Console.WriteLine(okWrong)
        Console.WriteLine(okRight)
        Console.WriteLine(cd("K"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "True", "99"]);
}

#[test]
fn test_vb_concurrent_dict_try_remove() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cd As New ConcurrentDictionary(Of Integer, String)()
        cd.TryAdd(1, "One")
        Dim removedVal As String = Nothing
        Dim ok As Boolean = cd.TryRemove(1, removedVal)
        Console.WriteLine(ok)
        Console.WriteLine(removedVal)
        Console.WriteLine(cd.IsEmpty)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "One", "True"]);
}

#[test]
fn test_vb_concurrent_dict_clear() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim cd As New ConcurrentDictionary(Of String, String)()
        cd.TryAdd("A", "B")
        cd.Clear()
        Console.WriteLine(cd.IsEmpty)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
