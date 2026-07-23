use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Dictionary(Of K, V) Key & Value Collections
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_dict_keys_collection_enumeration() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 1}, {"B", 2}, {"C", 3}}
        Dim keys As Dictionary(Of String, Integer).KeyCollection = dict.Keys
        Console.WriteLine(keys.Count)
        Console.WriteLine(String.Join(",", keys))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "A,B,C"]);
}

#[test]
fn test_vb_dict_values_collection_enumeration() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 10}, {"B", 20}, {"C", 30}}
        Dim values As Dictionary(Of String, Integer).ValueCollection = dict.Values
        Console.WriteLine(values.Count)
        Console.WriteLine(String.Join(",", values))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "10,20,30"]);
}

#[test]
fn test_vb_dict_custom_iequality_comparer() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        dict("foo") = 100
        Console.WriteLine(dict.ContainsKey("FOO"))
        Console.WriteLine(dict("FOO"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "100"]);
}

#[test]
fn test_vb_dict_try_add_method() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, String)
        Dim firstAdd As Boolean = dict.TryAdd(1, "First")
        Dim secondAdd As Boolean = dict.TryAdd(1, "Second")
        Console.WriteLine(firstAdd)
        Console.WriteLine(secondAdd)
        Console.WriteLine(dict(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False", "First"]);
}

#[test]
fn test_vb_dict_remove_key_value_pair() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, String) From {{1, "One"}, {2, "Two"}}
        Dim removedWrong As Boolean = dict.Remove(New KeyValuePair(Of Integer, String)(1, "Wrong"))
        Dim removedRight As Boolean = dict.Remove(New KeyValuePair(Of Integer, String)(1, "One"))
        Console.WriteLine(removedWrong)
        Console.WriteLine(removedRight)
        Console.WriteLine(dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "True", "1"]);
}

#[test]
fn test_vb_dict_ensure_capacity() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, String)()
        Dim cap As Integer = dict.EnsureCapacity(50)
        Console.WriteLine(cap >= 50)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_dict_trim_excess() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, Integer)(100)
        dict.Add(1, 10)
        dict.TrimExcess()
        Console.WriteLine(dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_dict_kvp_deconstruction() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"Key1", 100}}
        For Each kvp In dict
            Console.WriteLine(kvp.Key & ":" & kvp.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Key1:100"]);
}

#[test]
fn test_vb_dict_get_value_or_default() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"Existing", 50}}
        Dim val1 As Integer
        dict.TryGetValue("Existing", val1)
        Dim val2 As Integer
        dict.TryGetValue("Missing", val2)
        Console.WriteLine(val1)
        Console.WriteLine(val2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50", "0"]);
}

#[test]
fn test_vb_dict_clear_empties_collection() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, Integer) From {{1, 10}, {2, 20}}
        dict.Clear()
        Console.WriteLine(dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}
