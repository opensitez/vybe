use super::helpers::run_vb;

#[test]
fn dictionary_add_inserts_key_value_pair() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("a", 1)
        map.Add("b", 2)
        Console.WriteLine(map.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn dictionary_indexer_set_replaces_existing() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map("x") = 1
        map("x") = 9
        Console.WriteLine(map("x"))
        Console.WriteLine(map.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["9", "1"]);
}

#[test]
fn dictionary_contains_key_absent_key_is_false() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("a", 1)
        Console.WriteLine(map.ContainsKey("z"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False"]);
}

#[test]
fn dictionary_contains_value_detects_existing() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("a", 42)
        map.Add("b", 7)
        Console.WriteLine(map.ContainsValue(42))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn dictionary_try_get_value_returns_true_and_value_on_hit() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("k", 5)
        Dim value As Integer
        Console.WriteLine(map.TryGetValue("k", value))
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "5"]);
}

#[test]
fn dictionary_try_get_value_returns_false_and_default_on_miss() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        Dim value As Integer
        Console.WriteLine(map.TryGetValue("nope", value))
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "0"]);
}

#[test]
fn dictionary_remove_deletes_entry_and_reduces_count() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("a", 1)
        map.Add("b", 2)
        Console.WriteLine(map.Remove("a"))
        Console.WriteLine(map.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "1"]);
}

#[test]
fn dictionary_remove_missing_key_returns_false() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("a", 1)
        Console.WriteLine(map.Remove("z"))
        Console.WriteLine(map.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "1"]);
}

#[test]
fn dictionary_clear_removes_all_entries() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("a", 1)
        map.Add("b", 2)
        map.Clear()
        Console.WriteLine(map.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn dictionary_foreach_over_key_value_pairs() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("x", 10)
        map.Add("y", 20)
        For Each pair As KeyValuePair(Of String, Integer) In map
            Console.WriteLine(pair.Key & ":" & pair.Value)
        Next
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["x:10", "y:20"]);
}

#[test]
fn dictionary_keys_collect_values_and_sum() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("x", 10)
        map.Add("y", 20)
        map.Add("z", 30)

        Dim total As Integer = 0
        For Each key As String In map.Keys
            total += map(key)
        Next

        Console.WriteLine(map.Keys.Count)
        Console.WriteLine(total)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "60"]);
}

#[test]
fn dictionary_tryadd_insertes_when_missing() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        Console.WriteLine(map.TryAdd("a", 1))
        Console.WriteLine(map.TryAdd("a", 2))
        Console.WriteLine(map("a"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False", "1"]);
}

#[test]
fn dictionary_duplicate_add_raises_key_error() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("a", 1)
        Try
            map.Add("a", 2)
            Console.WriteLine("Added")
        Catch ex As ArgumentException
            Console.WriteLine("Duplicate")
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["Duplicate"]);
}

#[test]
fn dictionary_initialize_from_object_initializer() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer) From {
            {"x", 10},
            {"y", 20}
        }
        Console.WriteLine(map.Count)
        Console.WriteLine(map("x"))
        Console.WriteLine(map("y"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "10", "20"]);
}
