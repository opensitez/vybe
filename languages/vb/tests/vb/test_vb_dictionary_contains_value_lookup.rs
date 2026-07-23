use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Dictionary(Of TKey, TValue).ContainsValue & Surface
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_dictionary_contains_value_found() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 100}, {"B", 200}}
        Console.WriteLine(dict.ContainsValue(200))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_dictionary_contains_value_not_found() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 100}, {"B", 200}}
        Console.WriteLine(dict.ContainsValue(999))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_dictionary_contains_value_reference_type() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, String) From {{1, "Alpha"}, {2, "Beta"}}
        Console.WriteLine(dict.ContainsValue("Beta"))
        Console.WriteLine(dict.ContainsValue("Gamma"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_dictionary_contains_value_null_reference() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, String) From {{1, "Alpha"}, {2, Nothing}}
        Console.WriteLine(dict.ContainsValue(Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_dictionary_contains_key_contains_value_contrast() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, String) From {{"Key1", "Val1"}, {"Key2", "Val2"}}
        Console.WriteLine(dict.ContainsKey("Key1"))
        Console.WriteLine(dict.ContainsKey("Val1"))
        Console.WriteLine(dict.ContainsValue("Val1"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False", "True"]);
}

#[test]
fn test_vb_dictionary_try_get_value_success() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"Score", 95}}
        Dim val As Integer
        Dim found As Boolean = dict.TryGetValue("Score", val)
        Console.WriteLine(found & ":" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:95"]);
}

#[test]
fn test_vb_dictionary_try_get_value_failure_defaults() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"Score", 95}}
        Dim val As Integer
        Dim found As Boolean = dict.TryGetValue("Missing", val)
        Console.WriteLine(found & ":" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False:0"]);
}

#[test]
fn test_vb_dictionary_item_indexer_getter_setter() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, String)()
        dict(10) = "Ten"
        dict(20) = "Twenty"
        dict(10) = "TEN_UPDATED"
        Console.WriteLine(dict(10) & "|" & dict(20))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TEN_UPDATED|Twenty"]);
}

#[test]
fn test_vb_dictionary_add_duplicate_key_throws_or_handled() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer)()
        dict.Add("Unique", 1)
        Try
            dict.Add("Unique", 2)
            Console.WriteLine("Added Duplicate")
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentException"]);
}

#[test]
fn test_vb_dictionary_remove_key_returns_bool() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"K1", 1}, {"K2", 2}}
        Dim r1 As Boolean = dict.Remove("K1")
        Dim r2 As Boolean = dict.Remove("K1")
        Console.WriteLine(r1 & "|" & r2 & "|Count=" & dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False|Count=1"]);
}

#[test]
fn test_vb_dictionary_remove_key_with_out_value_parameter() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"Key", 42}}
        Dim removedVal As Integer
        Dim removed As Boolean = dict.Remove("Key", removedVal)
        Console.WriteLine(removed & ":" & removedVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:42"]);
}

#[test]
fn test_vb_dictionary_clear_resets_size() {
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

#[test]
fn test_vb_dictionary_values_collection_iteration() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 10}, {"B", 20}}
        Dim sum As Integer = 0
        For Each val In dict.Values
            sum += val
        Next
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30"]);
}

#[test]
fn test_vb_dictionary_keys_collection_iteration() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"X", 1}, {"Y", 2}}
        Dim keys As String = String.Join(",", dict.Keys)
        Console.WriteLine(keys)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["X,Y"]);
}

#[test]
fn test_vb_dictionary_struct_values_contains_value() {
    let src = r#"
Imports System.Collections.Generic

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Point) From {
            {"P1", New Point(1, 2)},
            {"P2", New Point(3, 4)}
        }
        Console.WriteLine(dict.ContainsValue(New Point(3, 4)))
        Console.WriteLine(dict.ContainsValue(New Point(0, 0)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_dictionary_custom_class_value_contains_value() {
    let src = r#"
Imports System.Collections.Generic

Class Element
    Public Property Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Overrides Function Equals(obj As Object) As Boolean
        Dim e = TryCast(obj, Element)
        Return e IsNot Nothing AndAlso e.Name = Me.Name
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return Name.GetHashCode()
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, Element) From {
            {1, New Element("Gold")},
            {2, New Element("Silver")}
        }
        Console.WriteLine(dict.ContainsValue(New Element("Gold")))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_dictionary_ensure_capacity_trim_excess() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, String)(100)
        dict(1) = "One"
        dict.TrimExcess()
        Console.WriteLine(dict.Count & "|" & dict(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|One"]);
}

#[test]
fn test_vb_dictionary_contains_value_multiple_duplicates() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 42}, {"B", 42}, {"C", 10}}
        Console.WriteLine(dict.ContainsValue(42))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_dictionary_enum_key_and_value() {
    let src = r#"
Imports System.Collections.Generic

Enum State
    Disabled
    Enabled
End Enum

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of State, State) From {{State.Disabled, State.Enabled}}
        Console.WriteLine(dict.ContainsKey(State.Disabled) & "|" & dict.ContainsValue(State.Enabled))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_dictionary_try_add_method() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer)()
        Dim added1 As Boolean = dict.TryAdd("Key", 100)
        Dim added2 As Boolean = dict.TryAdd("Key", 200)
        Console.WriteLine(added1 & "|" & added2 & "|" & dict("Key"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False|100"]);
}
