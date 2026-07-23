use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Collections.Generic.SortedDictionary Keys Ordering
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_sorted_dictionary_keys_ordered_automatically() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict(30) = "Thirty"
        dict(10) = "Ten"
        dict(20) = "Twenty"

        Dim keys = String.Join(",", dict.Keys)
        Console.WriteLine(keys)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_sorted_dictionary_custom_key_comparer_descending() {
    let src = r#"
Imports System.Collections.Generic

Class DescendingComparer
    Implements IComparer(Of Integer)
    Public Function Compare(x As Integer, y As Integer) As Integer Implements IComparer(Of Integer).Compare
        Return y.CompareTo(x)
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)(New DescendingComparer())
        dict(10) = "Ten"
        dict(30) = "Thirty"
        dict(20) = "Twenty"

        Dim keys = String.Join(",", dict.Keys)
        Console.WriteLine(keys)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30,20,10"]);
}

#[test]
fn test_vb_sorted_dictionary_string_keys_lexicographical() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)()
        dict("Banana") = 2
        dict("Apple") = 1
        dict("Cherry") = 3

        Dim keys = String.Join(",", dict.Keys)
        Console.WriteLine(keys)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Apple,Banana,Cherry"]);
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
        Dim found = dict.TryGetValue("Score", val)
        Dim missing = dict.TryGetValue("Missing", val)

        Console.WriteLine(found & "|" & val & "|" & missing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|100|False"]);
}

#[test]
fn test_vb_sorted_dictionary_remove_key() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict(1) = "One"
        dict(2) = "Two"
        dict(3) = "Three"

        Dim removed = dict.Remove(2)
        Dim remainingKeys = String.Join(",", dict.Keys)
        Console.WriteLine(removed & "|" & remainingKeys)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|1,3"]);
}

#[test]
fn test_vb_sorted_dictionary_contains_key_and_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, String)()
        dict("K1") = "V1"
        Console.WriteLine(dict.ContainsKey("K1") & "|" & dict.ContainsValue("V1") & "|" & dict.ContainsValue("V2"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|False"]);
}

#[test]
fn test_vb_sorted_dictionary_clear() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, Integer)()
        dict(1) = 100
        dict(2) = 200
        dict.Clear()
        Console.WriteLine(dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_sorted_dictionary_constructor_existing_dictionary() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim rawDict As New Dictionary(Of Integer, String)()
        rawDict(3) = "C"
        rawDict(1) = "A"
        rawDict(2) = "B"

        Dim sorted As New SortedDictionary(Of Integer, String)(rawDict)
        Console.WriteLine(String.Join(",", sorted.Keys))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_sorted_dictionary_custom_class_key_icomparable() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Class EmployeeKey
    Implements IComparable(Of EmployeeKey)
    Public Id As Integer
    Public Sub New(i As Integer)
        Id = i
    End Sub
    Public Function CompareTo(other As EmployeeKey) As Integer Implements IComparable(Of EmployeeKey).CompareTo
        Return Id.CompareTo(other.Id)
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of EmployeeKey, String)()
        dict(New EmployeeKey(50)) = "Fifty"
        dict(New EmployeeKey(10)) = "Ten"

        Dim ids As New List(Of Integer)()
        For Each k In dict.Keys
            ids.Add(k.Id)
        Next
        Console.WriteLine(String.Join(",", ids))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,50"]);
}

#[test]
fn test_vb_sorted_dictionary_values_ordered_by_keys() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict(3) = "Third"
        dict(1) = "First"
        dict(2) = "Second"

        Dim values = String.Join(",", dict.Values)
        Console.WriteLine(values)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["First,Second,Third"]);
}

#[test]
fn test_vb_sorted_dictionary_null_key_throws_argument_null() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)()
        Try
            dict.Add(Nothing, 100)
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
fn test_vb_sorted_dictionary_duplicate_key_add_throws_argument_exception() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)()
        dict.Add("Unique", 1)
        Try
            dict.Add("Unique", 2)
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException Caught on Duplicate Key Add")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentException Caught on Duplicate Key Add"]
    );
}

#[test]
fn test_vb_sorted_dictionary_indexer_assignment_updates_existing() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict(1) = "Initial"
        dict(1) = "Updated"
        Console.WriteLine(dict(1) & "|Count=" & dict.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Updated|Count=1"]);
}

#[test]
fn test_vb_sorted_dictionary_case_insensitive_string_comparer() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        dict("abc") = 10
        dict("XYZ") = 30
        dict("DEF") = 20

        Console.WriteLine(String.Join(",", dict.Keys))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["abc,DEF,XYZ"]);
}

#[test]
fn test_vb_sorted_dictionary_struct_value_type() {
    let src = r#"
Imports System.Collections.Generic

Structure ColorRGB
    Public R, G, B As Byte
End Structure

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, ColorRGB)()
        dict("Red") = New ColorRGB With {.R = 255, .G = 0, .B = 0}
        dict("Green") = New ColorRGB With {.R = 0, .G = 255, .B = 0}

        Console.WriteLine(dict("Green").G)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255"]);
}

#[test]
fn test_vb_sorted_dictionary_enumeration_kvp_order() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict(2) = "B"
        dict(1) = "A"

        Dim log = ""
        For Each kvp In dict
            log &= kvp.Key & ":" & kvp.Value & ";"
        Next
        Console.WriteLine(log)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:A;2:B;"]);
}

#[test]
fn test_vb_sorted_dictionary_linq_query_projections() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, Integer)()
        dict(1) = 10
        dict(2) = 20
        dict(3) = 30

        Dim query = dict.Where(Function(kvp) kvp.Value > 15).Select(Function(kvp) kvp.Key)
        Console.WriteLine(String.Join(",", query))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2,3"]);
}

#[test]
fn test_vb_sorted_dictionary_key_not_found_throws() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)()
        Try
            Dim val = dict(999)
        Catch ex As KeyNotFoundException
            Console.WriteLine("KeyNotFoundException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["KeyNotFoundException Caught"]);
}

#[test]
fn test_vb_sorted_dictionary_date_time_keys_order() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of DateTime, String)()
        dict(New DateTime(2025, 12, 31)) = "New Year Eve"
        dict(New DateTime(2025, 1, 1)) = "New Year Day"
        dict(New DateTime(2025, 6, 15)) = "Mid Year"

        Dim dates As New List(Of String)()
        For Each d In dict.Keys
            dates.Add(d.ToString("yyyy-MM-dd"))
        Next
        Console.WriteLine(String.Join(",", dates))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-01-01,2025-06-15,2025-12-31"]);
}

#[test]
fn test_vb_sorted_dictionary_copy_to_array_kvp() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)()
        dict("X") = 10
        dict("Y") = 20

        Dim array(1) As KeyValuePair(Of String, Integer)
        dict.CopyTo(array, 0)
        Console.WriteLine(array(0).Key & "=" & array(0).Value & "|" & array(1).Key & "=" & array(1).Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["X=10|Y=20"]);
}
