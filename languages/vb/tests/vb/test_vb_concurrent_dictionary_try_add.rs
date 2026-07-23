use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Collections.Concurrent.ConcurrentDictionary Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_concurrent_dictionary_try_add_and_try_getvalue() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        Dim added = dict.TryAdd("Key1", 100)
        Dim dupAdd = dict.TryAdd("Key1", 200)

        Dim val As Integer
        Dim found = dict.TryGetValue("Key1", val)
        Console.WriteLine(added & "|" & dupAdd & "|" & found & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False|True|100"]);
}

#[test]
fn test_vb_concurrent_dictionary_add_or_update_factory() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        ' Add initial value 10
        Dim v1 = dict.AddOrUpdate("Counter", 10, Function(key, oldVal) oldVal + 1)
        ' Update existing value
        Dim v2 = dict.AddOrUpdate("Counter", 10, Function(key, oldVal) oldVal + 1)
        Console.WriteLine(v1 & "|" & v2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10|11"]);
}

#[test]
fn test_vb_concurrent_dictionary_get_or_add_factory() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, String)()
        Dim s1 = dict.GetOrAdd("User1", Function(k) "NewUser_" & k)
        Dim s2 = dict.GetOrAdd("User1", Function(k) "OtherUser")
        Console.WriteLine(s1 & "|" & s2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NewUser_User1|NewUser_User1"]);
}

#[test]
fn test_vb_concurrent_dictionary_try_remove() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        dict("A") = 1

        Dim removedVal As Integer
        Dim ok = dict.TryRemove("A", removedVal)
        Dim okMissing = dict.TryRemove("B", removedVal)
        Console.WriteLine(ok & "|" & removedVal & "|" & okMissing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|1|False"]);
}

#[test]
fn test_vb_concurrent_dictionary_try_update_conditional() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        dict("State") = 100

        ' TryUpdate(key, newValue, comparisonValue)
        Dim okWrong = dict.TryUpdate("State", 200, 999) ' Comparison mismatch!
        Dim okRight = dict.TryUpdate("State", 200, 100) ' Comparison match!
        Console.WriteLine(okWrong & "|" & okRight & "|" & dict("State"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|True|200"]);
}

#[test]
fn test_vb_concurrent_dictionary_multithreaded_parallel_add() {
    let src = r#"
Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of Integer, Integer)()
        Parallel.For(0, 100, Sub(i) dict.TryAdd(i, i * 2))
        Console.WriteLine("Total Count: " & dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Total Count: 100"]);
}

#[test]
fn test_vb_concurrent_dictionary_clear() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        dict("X") = 1
        dict("Y") = 2
        dict.Clear()
        Console.WriteLine(dict.IsEmpty & "|" & dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|0"]);
}

#[test]
fn test_vb_concurrent_dictionary_custom_equality_comparer() {
    let src = r#"
Imports System
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        dict.TryAdd("HELLO", 42)
        Dim val As Integer
        Dim found = dict.TryGetValue("hello", val)
        Console.WriteLine(found & "|" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|42"]);
}

#[test]
fn test_vb_concurrent_dictionary_indexer_get_set() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, String)()
        dict("Prop") = "Val1"
        dict("Prop") = "Val2" ' Overwrites existing entry
        Console.WriteLine(dict("Prop"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Val2"]);
}

#[test]
fn test_vb_concurrent_dictionary_keys_and_values_snapshots() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        dict("A") = 1
        dict("B") = 2

        Dim keys = String.Join(",", dict.Keys)
        Dim values = String.Join(",", dict.Values)
        Console.WriteLine(dict.Keys.Count & "|" & dict.Values.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|2"]);
}

#[test]
fn test_vb_concurrent_dictionary_to_array_snapshot() {
    let src = r#"
Imports System.Collections.Concurrent
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        dict("K1") = 10
        Dim arr As KeyValuePair(Of String, Integer)() = dict.ToArray()
        Console.WriteLine(arr.Length & "|" & arr(0).Key)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|K1"]);
}

#[test]
fn test_vb_concurrent_dictionary_contains_key() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        dict.TryAdd("TargetKey", 99)
        Console.WriteLine(dict.ContainsKey("TargetKey") & "|" & dict.ContainsKey("MissingKey"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_concurrent_dictionary_concurrency_level_capacity_constructor() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of Integer, Integer)(concurrencyLevel := 4, capacity := 100)
        dict.TryAdd(1, 100)
        Console.WriteLine(dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_concurrent_dictionary_get_or_add_value_directly() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        Dim val1 = dict.GetOrAdd("FixedKey", 50)
        Dim val2 = dict.GetOrAdd("FixedKey", 100)
        Console.WriteLine(val1 & "|" & val2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50|50"]);
}

#[test]
fn test_vb_concurrent_dictionary_struct_key_and_value() {
    let src = r#"
Imports System.Collections.Concurrent

Structure Point2D
    Public X, Y As Integer
End Structure

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of Point2D, String)()
        Dim p As New Point2D With {.X = 1, .Y = 2}
        dict.TryAdd(p, "Point(1,2)")
        Console.WriteLine(dict(p))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Point(1,2)"]);
}

#[test]
fn test_vb_concurrent_dictionary_enumeration_during_mutation_safe() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of Integer, String)()
        dict(1) = "One"
        dict(2) = "Two"

        Dim count = 0
        For Each kvp In dict
            count += 1
            dict(count + 10) = "Extra" ' Safe to mutate during enumeration!
        Next
        Console.WriteLine(count >= 2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_concurrent_dictionary_null_key_throws_argument_null() {
    let src = r#"
Imports System
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        Try
            dict.TryAdd(Nothing, 10)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null Key")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null Key"]
    );
}

#[test]
fn test_vb_concurrent_dictionary_get_value_missing_key_throws_key_not_found() {
    let src = r#"
Imports System
Imports System.Collections.Generic
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        Try
            Dim val = dict("NonExistent")
        Catch ex As KeyNotFoundException
            Console.WriteLine("KeyNotFoundException Caught on Indexer Access")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["KeyNotFoundException Caught on Indexer Access"]
    );
}

#[test]
fn test_vb_concurrent_dictionary_add_or_update_with_arg_factory() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, Integer)()
        ' AddOrUpdate with factory parameter arg
        Dim res = dict.AddOrUpdate("Key", Function(k, arg) arg * 10, Function(k, oldVal, arg) oldVal + arg, 5)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50"]);
}

#[test]
fn test_vb_concurrent_dictionary_idictionary_explicit_implementation() {
    let src = r#"
Imports System.Collections
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim dict As IDictionary = New ConcurrentDictionary(Of String, Integer)()
        dict.Add("A", 100)
        Console.WriteLine(dict.Contains("A") & "|" & dict("A"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|100"]);
}
