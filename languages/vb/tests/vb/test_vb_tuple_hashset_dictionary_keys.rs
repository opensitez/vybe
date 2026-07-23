use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Tuples as HashSet Elements & Dictionary Keys
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_tuple_hashset_deduplication() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set As New HashSet(Of (Integer, Integer))()
        set.Add((1, 2))
        set.Add((1, 2))
        set.Add((2, 3))
        Console.WriteLine(set.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_tuple_dictionary_key_lookup() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (String, String), Integer)()
        dict(("US", "NY")) = 8000000
        dict(("US", "CA")) = 39000000
        Console.WriteLine(dict(("US", "NY")))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8000000"]);
}

#[test]
fn test_vb_tuple_dictionary_try_get_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (Integer, Integer), String)()
        dict((10, 20)) = "PointA"
        Dim res As String = Nothing
        Dim found = dict.TryGetValue((10, 20), res)
        Console.WriteLine(found & ":" & res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:PointA"]);
}

#[test]
fn test_vb_tuple_dictionary_contains_key() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (String, Integer), Boolean)()
        dict(("Admin", 1)) = True
        Console.WriteLine(dict.ContainsKey(("Admin", 1)) & "|" & dict.ContainsKey(("Guest", 2)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_named_tuple_dictionary_keys_name_erasure() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (X As Integer, Y As Integer), String)()
        dict((10, 20)) = "Location1"

        Dim searchTuple As (Col As Integer, Row As Integer) = (10, 20)
        Console.WriteLine(dict(searchTuple))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Location1"]);
}

#[test]
fn test_vb_tuple_hashset_union_with() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set1 As New HashSet(Of (Integer, String)) From {(1, "A"), (2, "B")}
        Dim set2 As New HashSet(Of (Integer, String)) From {(2, "B"), (3, "C")}
        set1.UnionWith(set2)
        Console.WriteLine(set1.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_tuple_hashset_intersect_with() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set1 As New HashSet(Of (Integer, String)) From {(1, "A"), (2, "B")}
        Dim set2 As New HashSet(Of (Integer, String)) From {(2, "B"), (3, "C")}
        set1.IntersectWith(set2)
        Console.WriteLine(set1.Count & ":" & set1.First().Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:B"]);
}

#[test]
fn test_vb_tuple_hashset_except_with() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set1 As New HashSet(Of (Integer, String)) From {(1, "A"), (2, "B")}
        Dim set2 As New HashSet(Of (Integer, String)) From {(2, "B")}
        set1.ExceptWith(set2)
        Console.WriteLine(set1.Count & ":" & set1.First().Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:A"]);
}

#[test]
fn test_vb_tuple_dictionary_remove_key() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (Integer, Integer), String) From {
            ((1, 1), "V1"),
            ((2, 2), "V2")
        }
        Dim removed = dict.Remove((1, 1))
        Console.WriteLine(removed & "|" & dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|1"]);
}

#[test]
fn test_vb_tuple_dictionary_clear() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (String, Integer), Double) From {{("A", 1), 10.0}}
        dict.Clear()
        Console.WriteLine(dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_tuple_dictionary_iteration() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (Integer, Integer), Integer)()
        dict((0, 0)) = 100
        dict((1, 1)) = 200

        For Each kv In dict
            Console.WriteLine(kv.Key.Item1 & "," & kv.Key.Item2 & "=" & kv.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0,0=100", "1,1=200"]);
}

#[test]
fn test_vb_tuple_sorted_dictionary_custom_key_ordering() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of (Integer, Integer), String)()
        dict((2, 1)) = "P21"
        dict((1, 5)) = "P15"
        dict((1, 2)) = "P12"

        For Each kv In dict
            Console.WriteLine(kv.Key.Item1 & "," & kv.Key.Item2 & "=" & kv.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2=P12", "1,5=P15", "2,1=P21"]);
}

#[test]
fn test_vb_tuple_sorted_set_ordering() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set As New SortedSet(Of (Integer, String)) From {
            (2, "B"),
            (1, "Z"),
            (1, "A")
        }
        For Each item In set
            Console.WriteLine(item.Item1 & ":" & item.Item2)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:A", "1:Z", "2:B"]);
}

#[test]
fn test_vb_tuple_hashset_contains_tuple() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set As New HashSet(Of (String, Integer)) From {("Alpha", 1), ("Beta", 2)}
        Console.WriteLine(set.Contains(("Alpha", 1)) & "|" & set.Contains(("Alpha", 2)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_tuple_dictionary_values_collection() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, (Name As String, Age As Integer)) From {
            {1, ("Alice", 25)},
            {2, ("Bob", 30)}
        }
        For Each val In dict.Values
            Console.WriteLine(val.Name & "=" & val.Age)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice=25", "Bob=30"]);
}

#[test]
fn test_vb_tuple_hashset_custom_iequalitycomparer() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Class TupleIgnoreCaseComparer
    Implements IEqualityComparer(Of (String, Integer))
    Public Function Equals(x As (String, Integer), y As (String, Integer)) As Boolean Implements IEqualityComparer(Of (String, Integer)).Equals
        Return String.Equals(x.Item1, y.Item1, StringComparison.OrdinalIgnoreCase) AndAlso x.Item2 = y.Item2
    End Function
    Public Function GetHashCode(obj As (String, Integer)) As Integer Implements IEqualityComparer(Of (String, Integer)).GetHashCode
        Return StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Item1) Xor obj.Item2.GetHashCode()
    End Function
End Class

Module Program
    Sub Main()
        Dim set As New HashSet(Of (String, Integer))(New TupleIgnoreCaseComparer())
        set.Add(("apple", 10))
        set.Add(("APPLE", 10))
        Console.WriteLine(set.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_tuple_dictionary_try_add() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (Integer, Integer), String)()
        Dim a1 = dict.TryAdd((1, 1), "V1")
        Dim a2 = dict.TryAdd((1, 1), "V2")
        Console.WriteLine(a1 & "|" & a2 & "|" & dict((1, 1)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False|V1"]);
}

#[test]
fn test_vb_tuple_dictionary_nested_tuple_values() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, (X As Integer, Y As Integer))()
        dict("P1") = (10, 20)
        dict("P2") = (30, 40)
        Console.WriteLine(dict("P1").X & "," & dict("P1").Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_tuple_hashset_array_conversion() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim set As New HashSet(Of (Integer, String)) From {(1, "A"), (2, "B")}
        Dim arr = set.ToArray()
        Console.WriteLine(arr.Length & ":" & arr(0).Item2 & "," & arr(1).Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2:A,B"]);
}

#[test]
fn test_vb_tuple_dictionary_lookup_null_element_in_tuple() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of (String, String), Integer)()
        dict((Nothing, "Val")) = 100
        Console.WriteLine(dict.ContainsKey((Nothing, "Val")) & "|" & dict((Nothing, "Val")))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|100"]);
}
