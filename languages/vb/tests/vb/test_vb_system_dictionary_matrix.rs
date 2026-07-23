use super::helpers::run_vb;

#[test]
fn dictionary_add_get_and_count() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("one", 1)
        map.Add("two", 2)
        Console.WriteLine(map.Count)
        Console.WriteLine(map("two"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn dictionary_update_via_indexer() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map("key") = 1
        map("key") = map("key") + 4
        Console.WriteLine(map("key"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5"]);
}

#[test]
fn dictionary_try_get_value_contract() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("a", 10)
        Dim value As Integer = 0
        Dim found As Boolean = map.TryGetValue("a", value)
        Dim missing As Integer = 0
        Dim missingFound As Boolean = map.TryGetValue("z", missing)
        Console.WriteLine(found)
        Console.WriteLine(value)
        Console.WriteLine(missingFound)
        Console.WriteLine(missing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "10", "False", "0"]);
}

#[test]
fn dictionary_contains_key() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, String)()
        map.Add("alpha", "A")
        Console.WriteLine(map.ContainsKey("alpha"))
        Console.WriteLine(map.ContainsKey("beta"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn dictionary_keys_values_counts() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map("a") = 1
        map("b") = 2
        map("c") = 3
        Console.WriteLine(map.Keys.Count)
        Console.WriteLine(map.Values.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn dictionary_remove_returns_flags() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of Integer, String)()
        map.Add(1, "one")
        map.Add(2, "two")
        Dim first As Boolean = map.Remove(1)
        Dim second As Boolean = map.Remove(3)
        Console.WriteLine(first)
        Console.WriteLine(second)
        Console.WriteLine(map.Count)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False", "1"]);
}

#[test]
fn dictionary_clear_resets_count() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, String)()
        map.Add("x", "X")
        map.Clear()
        Console.WriteLine(map.Count)
        Console.WriteLine(map.Count = 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0", "True"]);
}

#[test]
fn sorted_dictionary_ordering_contract() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New SortedDictionary(Of Integer, String)()
        map.Add(3, "three")
        map.Add(1, "one")
        map.Add(2, "two")
        Dim first As Integer = 0
        Dim second As Integer = 0
        Dim third As Integer = 0
        Dim i As Integer = 0
        For Each value As Integer In map.Keys
            If i = 0 Then
                first = value
            ElseIf i = 1 Then
                second = value
            ElseIf i = 2 Then
                third = value
            End If
            i += 1
        Next
        Console.WriteLine(first)
        Console.WriteLine(second)
        Console.WriteLine(third)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn dictionary_iteration_sum_values() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map("a") = 10
        map("b") = 20
        map("c") = 30
        Dim total As Integer = 0
        For Each kv As KeyValuePair(Of String, Integer) In map
            total += kv.Value
        Next
        Console.WriteLine(total)
        Console.WriteLine(map.Count)
    End Module
End Module
"#,
    );

    assert_eq!(out, vec!["60", "3"]);
}

#[test]
fn dictionary_contains_value_lookup() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of Integer, String)()
        map.Add(1, "yes")
        map.Add(2, "no")
        Console.WriteLine(map.ContainsValue("yes"))
        Console.WriteLine(map.ContainsValue("maybe"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}
