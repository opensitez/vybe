use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: SortedDictionary(Of TKey, TValue) & Custom Comparers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_sorted_dictionary_default_key_ordering() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict(30) = "Thirty"
        dict(10) = "Ten"
        dict(20) = "Twenty"
        Console.WriteLine(String.Join(",", dict.Keys))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_sorted_dictionary_string_key_alphabetical_ordering() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)()
        dict("Banana") = 2
        dict("Apple") = 1
        dict("Cherry") = 3
        Console.WriteLine(String.Join(",", dict.Keys))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Apple,Banana,Cherry"]);
}

#[test]
fn test_vb_sorted_dictionary_custom_reverse_comparer() {
    let src = r#"
Imports System.Collections.Generic

Class ReverseIntComparer
    Implements IComparer(Of Integer)
    Public Function Compare(x As Integer, y As Integer) As Integer Implements IComparer(Of Integer).Compare
        Return y.CompareTo(x)
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)(New ReverseIntComparer())
        dict(10) = "A" : dict(30) = "B" : dict(20) = "C"
        Console.WriteLine(String.Join(",", dict.Keys))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30,20,10"]);
}

#[test]
fn test_vb_sorted_dictionary_case_insensitive_string_comparer() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        dict("apple") = 1
        dict("BANANA") = 2
        dict("Cherry") = 3
        Console.WriteLine(dict.ContainsKey("APPLE") & "|" & dict("banana"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|2"]);
}

#[test]
fn test_vb_sorted_dictionary_try_get_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)()
        dict("Score") = 100
        Dim val As Integer
        Dim found As Boolean = dict.TryGetValue("Score", val)
        Console.WriteLine(found & ":" & val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:100"]);
}

#[test]
fn test_vb_sorted_dictionary_remove_key() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict(1) = "One" : dict(2) = "Two" : dict(3) = "Three"
        dict.Remove(2)
        Console.WriteLine(String.Join(",", dict.Keys))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,3"]);
}

#[test]
fn test_vb_sorted_dictionary_contains_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, String)()
        dict("K1") = "V1" : dict("K2") = "V2"
        Console.WriteLine(dict.ContainsValue("V2") & "|" & dict.ContainsValue("Missing"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_sorted_dictionary_custom_length_comparer() {
    let src = r#"
Imports System.Collections.Generic

Class StringLengthComparer
    Implements IComparer(Of String)
    Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
        Dim res = x.Length.CompareTo(y.Length)
        If res = 0 Then Return x.CompareTo(y)
        Return res
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)(New StringLengthComparer())
        dict("elephant") = 8
        dict("cat") = 3
        dict("dog") = 3
        Console.WriteLine(String.Join(",", dict.Keys))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["cat,dog,elephant"]);
}

#[test]
fn test_vb_sorted_dictionary_clear() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, Integer) From {{1, 10}, {2, 20}}
        dict.Clear()
        Console.WriteLine(dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_sorted_dictionary_key_value_pair_iteration() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer) From {{"B", 2}, {"A", 1}}
        For Each kv In dict
            Console.WriteLine(kv.Key & "=" & kv.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A=1", "B=2"]);
}

#[test]
fn test_vb_sorted_dictionary_struct_keys_icomparable() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Structure Point
    Implements IComparable(Of Point)
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
    Public Function CompareTo(other As Point) As Integer Implements IComparable(Of Point).CompareTo
        Dim res = X.CompareTo(other.X)
        If res = 0 Then Return Y.CompareTo(other.Y)
        Return res
    End Function
End Structure

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Point, String)()
        dict(New Point(2, 1)) = "P21"
        dict(New Point(1, 5)) = "P15"
        dict(New Point(1, 2)) = "P12"
        For Each kv In dict
            Console.WriteLine(kv.Key.X & "," & kv.Key.Y & "=" & kv.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2=P12", "1,5=P15", "2,1=P21"]);
}

#[test]
fn test_vb_sorted_dictionary_values_ordered_by_keys() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict(3) = "C" : dict(1) = "A" : dict(2) = "B"
        Console.WriteLine(String.Join(",", dict.Values))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A,B,C"]);
}

#[test]
fn test_vb_sorted_dictionary_enum_keys() {
    let src = r#"
Imports System.Collections.Generic

Enum Priority
    Low = 0
    Medium = 1
    High = 2
End Enum

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Priority, String)()
        dict(Priority.High) = "Emergency"
        dict(Priority.Low) = "Routine"
        For Each kv In dict
            Console.WriteLine(kv.Key.ToString() & ":" & kv.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Low:Routine", "High:Emergency"]);
}

#[test]
fn test_vb_sorted_dictionary_comparer_property() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Console.WriteLine(dict.Comparer Is StringComparer.OrdinalIgnoreCase)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_sorted_dictionary_copy_to_key_value_pairs() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String) From {{2, "Two"}, {1, "One"}}
        Dim arr(1) As KeyValuePair(Of Integer, String)
        dict.CopyTo(arr, 0)
        Console.WriteLine(arr(0).Key & "=" & arr(0).Value)
        Console.WriteLine(arr(1).Key & "=" & arr(1).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1=One", "2=Two"]);
}

#[test]
fn test_vb_sorted_dictionary_duplicate_key_exception() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict.Add(1, "One")
        Try
            dict.Add(1, "Duplicate")
        Catch ex As ArgumentException
            Console.WriteLine("DuplicateKeyException")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["DuplicateKeyException"]);
}

#[test]
fn test_vb_sorted_dictionary_tuple_keys() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of (Integer, Integer), String)()
        dict((2, 1)) = "B"
        dict((1, 5)) = "A"
        For Each kv In dict
            Console.WriteLine(kv.Key.Item1 & "," & kv.Key.Item2 & ":" & kv.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,5:A", "2,1:B"]);
}

#[test]
fn test_vb_sorted_dictionary_datetime_keys() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of DateTime, String)()
        dict(New DateTime(2025, 12, 31)) = "New Year's Eve"
        dict(New DateTime(2025, 1, 1)) = "New Year's Day"
        For Each kv In dict
            Console.WriteLine(kv.Key.ToString("yyyy-MM-dd") & "=" & kv.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["2025-01-01=New Year's Day", "2025-12-31=New Year's Eve"]
    );
}

#[test]
fn test_vb_sorted_dictionary_update_existing_key() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer) From {{"K1", 10}}
        dict("K1") = 99
        Console.WriteLine(dict("K1"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99"]);
}

#[test]
fn test_vb_sorted_dictionary_empty_dictionary() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, String)()
        Console.WriteLine(dict.Count & "|" & (dict.Keys.Count = 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0|True"]);
}
