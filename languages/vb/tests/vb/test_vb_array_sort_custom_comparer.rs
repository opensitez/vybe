use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array Sort with Custom Comparers & Keys
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_sort_primitive_ascending() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {5, 2, 8, 1, 9}
        Array.Sort(arr)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,5,8,9"]);
}

#[test]
fn test_vb_array_sort_keys_and_items() {
    let src = r#"
Module Program
    Sub Main()
        Dim keys As Integer() = {3, 1, 2}
        Dim items As String() = {"Three", "One", "Two"}
        Array.Sort(keys, items)
        Console.WriteLine(String.Join(",", items))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["One,Two,Three"]);
}

#[test]
fn test_vb_array_sort_custom_icomparer_generic() {
    let src = r#"
Imports System.Collections.Generic

Class LengthComparer
    Implements IComparer(Of String)
    Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
        Return x.Length.CompareTo(y.Length)
    End Function
End Class

Module Program
    Sub Main()
        Dim words As String() = {"Elephant", "Cat", "Giraffe", "Dog"}
        Array.Sort(words, New LengthComparer())
        Console.WriteLine(String.Join(",", words))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Cat,Dog,Giraffe,Elephant"]);
}

#[test]
fn test_vb_array_sort_comparison_delegate() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {1, 5, 2, 9, 3}
        Array.Sort(arr, Function(x, y) y.CompareTo(x))
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9,5,3,2,1"]);
}

#[test]
fn test_vb_array_sort_range_index_count() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {5, 4, 3, 2, 1}
        Array.Sort(arr, 1, 3)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5,2,3,4,1"]);
}

#[test]
fn test_vb_array_sort_icomparable_objects() {
    let src = r#"
Class Person
    Implements IComparable(Of Person)
    Public Property Age As Integer
    Public Property Name As String

    Public Function CompareTo(other As Person) As Integer Implements IComparable(Of Person).CompareTo
        Return Age.CompareTo(other.Age)
    End Function
End Class

Module Program
    Sub Main()
        Dim people As Person() = {
            New Person With {.Age = 30, .Name = "Bob"},
            New Person With {.Age = 20, .Name = "Alice"}
        }
        Array.Sort(people)
        Console.WriteLine(people(0).Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice"]);
}

#[test]
fn test_vb_array_sort_string_case_insensitive() {
    let src = r#"
Imports System.Collections

Module Program
    Sub Main()
        Dim words As String() = {"b", "A", "c", "B"}
        Array.Sort(words, StringComparer.OrdinalIgnoreCase)
        Console.WriteLine(String.Join(",", words))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A,b,B,c"]);
}

#[test]
fn test_vb_array_sort_double_keys_string_items() {
    let src = r#"
Module Program
    Sub Main()
        Dim scores As Double() = {98.5, 87.0, 92.3}
        Dim students As String() = {"Alice", "Bob", "Charlie"}
        Array.Sort(scores, students)
        Console.WriteLine(students(0))
        Console.WriteLine(students(2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Bob", "Alice"]);
}

#[test]
fn test_vb_array_sort_stable_order_preservation() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 10, 30}
        Array.Sort(arr)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,10,20,30"]);
}

#[test]
fn test_vb_array_sort_range_keys_and_items() {
    let src = r#"
Module Program
    Sub Main()
        Dim keys As Integer() = {10, 40, 30, 20, 50}
        Dim vals As String() = {"A", "D", "C", "B", "E"}
        Array.Sort(keys, vals, 1, 3)
        Console.WriteLine(String.Join(",", vals))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A,B,C,D,E"]);
}
