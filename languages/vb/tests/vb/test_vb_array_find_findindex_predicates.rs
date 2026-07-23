use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Array.Find, FindIndex, FindLast, FindAll & Predicates
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_find_first_matching_element() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 3, 5, 8, 10, 12}
        Dim even As Integer = Array.Find(numbers, Function(n) n Mod 2 = 0)
        Console.WriteLine(even)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_array_find_no_match_default() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 3, 5, 7}
        Dim even As Integer = Array.Find(numbers, Function(n) n Mod 2 = 0)
        Console.WriteLine(even)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_array_findindex_first_matching_index() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim words As String() = {"cat", "elephant", "dog"}
        Dim idx As Integer = Array.FindIndex(words, Function(w) w.Length > 5)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_array_findindex_no_match_minus_one() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim words As String() = {"cat", "dog"}
        Dim idx As Integer = Array.FindIndex(words, Function(w) w.Length > 5)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1"]);
}

#[test]
fn test_vb_array_findlast_matching_element() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {2, 4, 7, 8, 10, 11}
        Dim lastEven As Integer = Array.FindLast(numbers, Function(n) n Mod 2 = 0)
        Console.WriteLine(lastEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_array_findlastindex_matching_index() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {2, 4, 7, 8, 10, 11}
        Dim lastEvenIdx As Integer = Array.FindLastIndex(numbers, Function(n) n Mod 2 = 0)
        Console.WriteLine(lastEvenIdx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}

#[test]
fn test_vb_array_findall_matching_elements_array() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4, 5, 6}
        Dim evens As Integer() = Array.FindAll(numbers, Function(n) n Mod 2 = 0)
        Console.WriteLine(String.Join(",", evens))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2,4,6"]);
}

#[test]
fn test_vb_array_findall_no_matches_empty_array() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 3, 5}
        Dim evens As Integer() = Array.FindAll(numbers, Function(n) n Mod 2 = 0)
        Console.WriteLine(evens.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_array_findindex_start_index_offset() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {2, 4, 6, 8, 10}
        ' Search from index 2 onwards
        Dim idx As Integer = Array.FindIndex(numbers, 2, Function(n) n > 5)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_array_findindex_start_index_and_count() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {2, 4, 6, 8, 10}
        ' Search range [1, 1+2] = indices 1 and 2 (values 4, 6)
        Dim idx As Integer = Array.FindIndex(numbers, 1, 2, Function(n) n > 5)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_array_findindex_start_index_and_count_no_match() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {2, 4, 6, 8, 10}
        ' Search range [0, 2] = indices 0 and 1 (values 2, 4)
        Dim idx As Integer = Array.FindIndex(numbers, 0, 2, Function(n) n > 5)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1"]);
}

#[test]
fn test_vb_array_find_with_complex_object_predicate() {
    let src = r#"
Imports System

Class Person
    Public Property Name As String
    Public Property Age As Integer
    Public Sub New(n As String, a As Integer)
        Name = n : Age = a
    End Sub
End Class

Module Program
    Sub Main()
        Dim people As Person() = {New Person("Alice", 25), New Person("Bob", 35), New Person("Charlie", 30)}
        Dim found As Person = Array.Find(people, Function(p) p.Age > 30)
        Console.WriteLine(found.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Bob"]);
}

#[test]
fn test_vb_array_findlast_with_complex_object() {
    let src = r#"
Imports System

Class Product
    Public Property Category As String
    Public Property Price As Double
    Public Sub New(c As String, p As Double)
        Category = c : Price = p
    End Sub
End Class

Module Program
    Sub Main()
        Dim prods As Product() = {
            New Product("Tech", 100),
            New Product("Food", 5),
            New Product("Tech", 500)
        }
        Dim lastTech As Product = Array.FindLast(prods, Function(p) p.Category = "Tech")
        Console.WriteLine(lastTech.Price)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["500"]);
}

#[test]
fn test_vb_array_find_string_case_insensitive_predicate() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim tags As String() = {"ALPHA", "BETA", "GAMMA"}
        Dim found As String = Array.Find(tags, Function(t) t.Equals("beta", StringComparison.OrdinalIgnoreCase))
        Console.WriteLine(found)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["BETA"]);
}

#[test]
fn test_vb_array_find_struct_predicate() {
    let src = r#"
Imports System

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim pts As Point() = {New Point(1, 2), New Point(3, 4), New Point(5, 6)}
        Dim target As Point = Array.Find(pts, Function(p) p.X = 3)
        Console.WriteLine(target.X & "," & target.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3,4"]);
}

#[test]
fn test_vb_array_find_predicate_side_effects_counter() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim calls As Integer = 0
        Dim numbers As Integer() = {10, 20, 30, 40, 50}
        Dim match As Integer = Array.Find(numbers, Function(n)
            calls += 1
            Return n = 30
        End Function)
        Console.WriteLine(match & "|calls=" & calls)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30|calls=3"]);
}

#[test]
fn test_vb_array_findlastindex_start_index_backwards() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30, 20, 10}
        ' Search backwards from index 2
        Dim idx As Integer = Array.FindLastIndex(numbers, 2, Function(n) n = 20)
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_array_find_nullable_types() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim items As Nullable(Of Integer)() = {10, Nothing, 30, Nothing}
        Dim firstNullIdx As Integer = Array.FindIndex(items, Function(i) Not i.HasValue)
        Console.WriteLine(firstNullIdx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_array_find_enum_values() {
    let src = r#"
Imports System

Enum Status
    Pending
    Active
    Inactive
End Enum

Module Program
    Sub Main()
        Dim list As Status() = {Status.Pending, Status.Active, Status.Inactive}
        Dim match As Status = Array.Find(list, Function(s) s = Status.Active)
        Console.WriteLine(match.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Active"]);
}

#[test]
fn test_vb_array_find_all_empty_predicate() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30}
        Dim allMatches As Integer() = Array.FindAll(numbers, Function(n) True)
        Console.WriteLine(allMatches.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}
