use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ Distinct, DistinctBy & Custom Comparers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_distinct_primitive_integers() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 2, 3, 1, 4, 3}
        Dim unique = numbers.Distinct()
        Console.WriteLine(String.Join(",", unique))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,4"]);
}

#[test]
fn test_vb_linq_distinct_strings_case_sensitive() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "APPLE", "banana", "apple"}
        Dim unique = words.Distinct()
        Console.WriteLine(String.Join(",", unique))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["apple,APPLE,banana"]);
}

#[test]
fn test_vb_linq_distinct_string_comparer_ignore_case() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "APPLE", "banana", "Apple"}
        Dim unique = words.Distinct(StringComparer.OrdinalIgnoreCase)
        Console.WriteLine(String.Join(",", unique))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["apple,banana"]);
}

#[test]
fn test_vb_linq_distinct_by_property_key_selector() {
    let src = r#"
Imports System.Linq

Class Person
    Public Property Name As String
    Public Property Age As Integer
    Public Sub New(n As String, a As Integer) : Name = n : Age = a : End Sub
End Class

Module Program
    Sub Main()
        Dim people = {New Person("Alice", 25), New Person("Bob", 25), New Person("Charlie", 30)}
        Dim uniqueByAge = people.DistinctBy(Function(p) p.Age)
        For Each p In uniqueByAge
            Console.WriteLine(p.Name & ":" & p.Age)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice:25", "Charlie:30"]);
}

#[test]
fn test_vb_linq_distinct_custom_iequalitycomparer() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Class Product
    Public Property ID As Integer
    Public Property Name As String
    Public Sub New(id As Integer, name As String) : Me.ID = id : Me.Name = name : End Sub
End Class

Class ProductIDComparer
    Implements IEqualityComparer(Of Product)
    Public Function Equals(x As Product, y As Product) As Boolean Implements IEqualityComparer(Of Product).Equals
        If x Is y Then Return True
        If x Is Nothing OrElse y Is Nothing Then Return False
        Return x.ID = y.ID
    End Function
    Public Function GetHashCode(obj As Product) As Integer Implements IEqualityComparer(Of Product).GetHashCode
        If obj Is Nothing Then Return 0
        Return obj.ID.GetHashCode()
    End Function
End Class

Module Program
    Sub Main()
        Dim prods = {New Product(1, "P1"), New Product(1, "P1_Dup"), New Product(2, "P2")}
        Dim unique = prods.Distinct(New ProductIDComparer())
        For Each p In unique
            Console.WriteLine(p.ID & "=" & p.Name)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1=P1", "2=P2"]);
}

#[test]
fn test_vb_linq_distinct_tuples() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim pairs = {("A", 1), ("B", 2), ("A", 1), ("A", 2)}
        Dim unique = pairs.Distinct()
        Console.WriteLine(unique.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_linq_distinct_structs() {
    let src = r#"
Imports System.Linq

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer) : Me.X = x : Me.Y = y : End Sub
End Structure

Module Program
    Sub Main()
        Dim points = {New Point(1, 2), New Point(1, 2), New Point(3, 4)}
        Dim unique = points.Distinct()
        Console.WriteLine(unique.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_linq_distinct_query_syntax() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 2, 3, 1}
        Dim query = From n In numbers Distinct
        Console.WriteLine(String.Join(",", query))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_linq_distinct_empty_sequence() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim empty As Integer() = {}
        Dim unique = empty.Distinct()
        Console.WriteLine(unique.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_linq_distinct_all_duplicates() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim same = {5, 5, 5, 5}
        Dim unique = same.Distinct()
        Console.WriteLine(String.Join(",", unique))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_linq_distinct_null_elements_handled() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim strings As String() = {"A", Nothing, "B", Nothing, "A"}
        Dim unique = strings.Distinct()
        Console.WriteLine(unique.Count() & "|" & (unique.Last() Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3|False"]);
}

#[test]
fn test_vb_linq_distinct_by_string_length() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"cat", "dog", "elephant", "bear", "fox"}
        Dim uniqueLengths = words.DistinctBy(Function(w) w.Length)
        Console.WriteLine(String.Join(",", uniqueLengths))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["cat,elephant,bear"]);
}

#[test]
fn test_vb_linq_distinct_by_custom_key_comparer() {
    let src = r#"
Imports System
Imports System.Linq

Class Employee
    Public Property Department As String
    Public Property Name As String
    Public Sub New(d As String, n As String) : Department = d : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim emps = {New Employee("hr", "Alice"), New Employee("HR", "Bob"), New Employee("IT", "Charlie")}
        Dim uniqueDepts = emps.DistinctBy(Function(e) e.Department, StringComparer.OrdinalIgnoreCase)
        For Each e In uniqueDepts
            Console.WriteLine(e.Department & ":" & e.Name)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["hr:Alice", "IT:Charlie"]);
}

#[test]
fn test_vb_linq_distinct_preserves_first_occurrence() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20, 10, 30, 20}
        Dim unique = numbers.Distinct()
        Console.WriteLine(String.Join(",", unique))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_linq_distinct_deferred_execution() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 2}
        Dim query = list.Distinct()
        list.Add(3)
        list.Add(3)
        Console.WriteLine(String.Join(",", query))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_linq_distinct_enum_values() {
    let src = r#"
Imports System.Linq

Enum Status
    Active
    Pending
End Enum

Module Program
    Sub Main()
        Dim statuses = {Status.Active, Status.Pending, Status.Active}
        Dim unique = statuses.Distinct()
        Console.WriteLine(unique.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_linq_distinct_datetime_dates_only() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim dates = {New DateTime(2025, 1, 1, 10, 0, 0), New DateTime(2025, 1, 1, 14, 0, 0), New DateTime(2025, 1, 2, 9, 0, 0)}
        Dim uniqueDays = dates.DistinctBy(Function(d) d.Date)
        Console.WriteLine(uniqueDays.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_linq_distinct_char_array() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim chars As Char() = {"a"c, "b"c, "a"c, "c"c}
        Dim unique = chars.Distinct()
        Console.WriteLine(New String(unique.ToArray()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["abc"]);
}

#[test]
fn test_vb_linq_distinct_anonymous_types() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim items = {New With {.ID = 1, .Val = "A"}, New With {.ID = 1, .Val = "A"}, New With {.ID = 2, .Val = "B"}}
        Dim unique = items.Distinct()
        Console.WriteLine(unique.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_linq_distinct_chained_with_where_select() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5, 6, 7, 8}
        ' Filter evens, divide by 2, get distinct
        Dim result = numbers.Where(Function(n) n Mod 2 = 0).Select(Function(n) n \ 2).Distinct()
        Console.WriteLine(String.Join(",", result))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,4"]);
}
