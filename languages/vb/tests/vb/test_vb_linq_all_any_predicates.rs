use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ All, Any & Sequence Matching Operators
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_any_no_args_non_empty() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3}
        Console.WriteLine(numbers.Any())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_any_no_args_empty() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim empty As Integer() = {}
        Console.WriteLine(empty.Any())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_linq_any_predicate_true() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 3, 5, 8}
        Dim hasEven = numbers.Any(Function(n) n Mod 2 = 0)
        Console.WriteLine(hasEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_any_predicate_false() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 3, 5, 7}
        Dim hasEven = numbers.Any(Function(n) n Mod 2 = 0)
        Console.WriteLine(hasEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_linq_all_predicate_true() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim evens = {2, 4, 6, 8}
        Dim allEven = evens.All(Function(n) n Mod 2 = 0)
        Console.WriteLine(allEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_all_predicate_false() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {2, 4, 5, 8}
        Dim allEven = numbers.All(Function(n) n Mod 2 = 0)
        Console.WriteLine(allEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_linq_all_empty_sequence_vacuously_true() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim empty As Integer() = {}
        Console.WriteLine(empty.All(Function(n) n > 100))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_any_short_circuits() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim count As Integer = 0
        Dim numbers = {10, -5, 20, 30}
        Dim hasNeg = numbers.Any(Function(n)
            count += 1
            Return n < 0
        End Function)
        Console.WriteLine(hasNeg & "|count=" & count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|count=2"]);
}

#[test]
fn test_vb_linq_all_short_circuits() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim count As Integer = 0
        Dim numbers = {10, -5, 20, 30}
        Dim allPos = numbers.All(Function(n)
            count += 1
            Return n > 0
        End Function)
        Console.WriteLine(allPos & "|count=" & count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|count=2"]);
}

#[test]
fn test_vb_linq_contains_primitive_item() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20, 30}
        Console.WriteLine(numbers.Contains(20) & "|" & numbers.Contains(99))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_linq_contains_string_custom_comparer() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "banana", "cherry"}
        Console.WriteLine(words.Contains("BANANA", StringComparer.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_contains_complex_object_reference() {
    let src = r#"
Imports System.Linq

Class Product
    Public Property Name As String
    Public Sub New(n As String) : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim p1 As New Product("Laptop")
        Dim p2 As New Product("Phone")
        Dim list = {p1}
        Console.WriteLine(list.Contains(p1) & "|" & list.Contains(p2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_linq_contains_struct_value_equality() {
    let src = r#"
Imports System.Linq

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer) : Me.X = x : Me.Y = y : End Sub
End Structure

Module Program
    Sub Main()
        Dim points = {New Point(1, 2), New Point(3, 4)}
        Console.WriteLine(points.Contains(New Point(3, 4)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_any_string_length_condition() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"cat", "elephant", "dog"}
        Console.WriteLine(words.Any(Function(w) w.Length > 5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_all_string_start_with_condition() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "apricot", "avocado"}
        Console.WriteLine(words.All(Function(w) w.StartsWith("a")))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_any_nested_in_where_filter() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Class Department
    Public Property Name As String
    Public Property Employees As List(Of String)
End Class

Module Program
    Sub Main()
        Dim depts As New List(Of Department) From {
            New Department With {.Name = "HR", .Employees = New List(Of String) From {"Alice"}},
            New Department With {.Name = "IT", .Employees = New List(Of String) From {"Bob", "Charlie"}}
        }
        Dim withCharlie = depts.Where(Function(d) d.Employees.Any(Function(e) e = "Charlie"))
        Console.WriteLine(withCharlie.First().Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IT"]);
}

#[test]
fn test_vb_linq_all_nested_in_where_filter() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Class Team
    Public Property Scores As List(Of Integer)
End Class

Module Program
    Sub Main()
        Dim teams As New List(Of Team) From {
            New Team With {.Scores = New List(Of Integer) From {10, 20, 30}},
            New Team With {.Scores = New List(Of Integer) From {25, 35, 45}}
        }
        Dim highScorers = teams.Where(Function(t) t.Scores.All(Function(s) s >= 20))
        Console.WriteLine(highScorers.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_linq_any_nullable_types_has_value() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim items As Nullable(Of Integer)() = {Nothing, 10, Nothing}
        Console.WriteLine(items.Any(Function(i) i.HasValue))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_all_enum_values_check() {
    let src = r#"
Imports System.Linq

Enum Status
    Active
    Pending
End Enum

Module Program
    Sub Main()
        Dim statuses = {Status.Active, Status.Active}
        Console.WriteLine(statuses.All(Function(s) s = Status.Active))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_any_dictionary_kvp_matching() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 10}, {"B", 20}}
        Console.WriteLine(dict.Any(Function(kv) kv.Value > 15))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
