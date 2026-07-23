use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String.Join Overloads & Full Collection Surface
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_string_join_array_of_string_comma() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As String() = {"Apple", "Banana", "Cherry"}
        Console.WriteLine(String.Join(", ", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Apple, Banana, Cherry"]);
}

#[test]
fn test_vb_string_join_array_of_object() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Object() = {1, "Two", 3.0, True}
        Console.WriteLine(String.Join(" | ", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1 | Two | 3 | True"]);
}

#[test]
fn test_vb_string_join_generic_list_of_int() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {10, 20, 30, 40}
        Console.WriteLine(String.Join("-", list))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10-20-30-40"]);
}

#[test]
fn test_vb_string_join_generic_enumerable_of_double() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim nums As IEnumerable(Of Double) = {1.1, 2.2, 3.3}
        Console.WriteLine(String.Join(" ; ", nums))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.1 ; 2.2 ; 3.3"]);
}

#[test]
fn test_vb_string_join_char_delimiter_overload() {
    let src = r#"
Module Program
    Sub Main()
        Dim words = {"foo", "bar", "baz"}
        Console.WriteLine(String.Join("/"c, words))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["foo/bar/baz"]);
}

#[test]
fn test_vb_string_join_array_start_index_and_count() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As String() = {"A", "B", "C", "D", "E"}
        Console.WriteLine(String.Join(":", arr, 1, 3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B:C:D"]);
}

#[test]
fn test_vb_string_join_array_start_index_zero_count_all() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As String() = {"X", "Y", "Z"}
        Console.WriteLine(String.Join(".", arr, 0, 3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["X.Y.Z"]);
}

#[test]
fn test_vb_string_join_empty_delimiter() {
    let src = r#"
Module Program
    Sub Main()
        Dim words = {"H", "e", "l", "l", "o"}
        Console.WriteLine(String.Join("", words))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello"]);
}

#[test]
fn test_vb_string_join_single_element() {
    let src = r#"
Module Program
    Sub Main()
        Dim singleItem = {"SoleElement"}
        Console.WriteLine(String.Join(", ", singleItem))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SoleElement"]);
}

#[test]
fn test_vb_string_join_empty_collection() {
    let src = r#"
Module Program
    Sub Main()
        Dim emptyArr As String() = {}
        Console.WriteLine("'" & String.Join(", ", emptyArr) & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["''"]);
}

#[test]
fn test_vb_string_join_null_elements_handled_as_empty() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As String() = {"One", Nothing, "Three"}
        Console.WriteLine(String.Join("-", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["One--Three"]);
}

#[test]
fn test_vb_string_join_multiline_delimiter() {
    let src = r#"
Module Program
    Sub Main()
        Dim lines = {"Header", "Body", "Footer"}
        Console.WriteLine(String.Join(vbCrLf, lines))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Header", "Body", "Footer"]);
}

#[test]
fn test_vb_string_join_char_delimiter_with_sub_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As String() = {"10", "20", "30", "40"}
        Console.WriteLine(String.Join(","c, arr, 1, 2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20,30"]);
}

#[test]
fn test_vb_string_join_linq_projection() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5}
        Dim result = String.Join(" + ", numbers.Select(Function(n) n * 2))
        Console.WriteLine(result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2 + 4 + 6 + 8 + 10"]);
}

#[test]
fn test_vb_string_join_struct_array() {
    let src = r#"
Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
    Public Overrides Function ToString() As String
        Return "(" & X & "," & Y & ")"
    End Function
End Structure

Module Program
    Sub Main()
        Dim pts As Point() = {New Point(1, 2), New Point(3, 4)}
        Console.WriteLine(String.Join("; ", pts))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["(1,2); (3,4)"]);
}

#[test]
fn test_vb_string_join_custom_class_enumerable() {
    let src = r#"
Imports System.Collections.Generic

Class Person
    Public Property Name As String
    Public Sub New(n As String)
        Me.Name = n
    End Sub
    Public Overrides Function ToString() As String
        Return Name
    End Function
End Class

Module Program
    Sub Main()
        Dim people As New List(Of Person) From {New Person("Alice"), New Person("Bob")}
        Console.WriteLine(String.Join(" & ", people))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice & Bob"]);
}

#[test]
fn test_vb_string_join_params_array_syntax() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(String.Join(", ", "Alpha", "Beta", "Gamma"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alpha, Beta, Gamma"]);
}

#[test]
fn test_vb_string_join_dictionary_entries() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 1}, {"B", 2}}
        Dim joined = String.Join(", ", dict.Select(Function(kv) kv.Key & "=" & kv.Value))
        Console.WriteLine(joined)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A=1, B=2"]);
}

#[test]
fn test_vb_string_join_long_delimiter() {
    let src = r#"
Module Program
    Sub Main()
        Dim parts = {"Part1", "Part2"}
        Console.WriteLine(String.Join(" ==> ", parts))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Part1 ==> Part2"]);
}

#[test]
fn test_vb_string_join_count_zero() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As String() = {"A", "B", "C"}
        Console.WriteLine("'" & String.Join(",", arr, 1, 0) & "'")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["''"]);
}
