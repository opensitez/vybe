use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ Union, Intersect, Except & Set Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_union_combines_and_removes_duplicates() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2, 3}
        Dim seq2 = {3, 4, 5}
        Dim res = seq1.Union(seq2)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,4,5"]);
}

#[test]
fn test_vb_linq_intersect_finds_common_elements() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2, 3, 4}
        Dim seq2 = {3, 4, 5, 6}
        Dim res = seq1.Intersect(seq2)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3,4"]);
}

#[test]
fn test_vb_linq_except_removes_elements_present_in_second() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2, 3, 4, 5}
        Dim seq2 = {2, 4}
        Dim res = seq1.Except(seq2)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,3,5"]);
}

#[test]
fn test_vb_linq_union_string_case_insensitive_comparer() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {"apple", "banana"}
        Dim seq2 = {"BANANA", "CHERRY"}
        Dim res = seq1.Union(seq2, StringComparer.OrdinalIgnoreCase)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["apple,banana,CHERRY"]);
}

#[test]
fn test_vb_linq_intersect_string_case_insensitive_comparer() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {"apple", "banana"}
        Dim seq2 = {"BANANA", "CHERRY"}
        Dim res = seq1.Intersect(seq2, StringComparer.OrdinalIgnoreCase)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["banana"]);
}

#[test]
fn test_vb_linq_except_string_case_insensitive_comparer() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {"APPLE", "BANANA", "CHERRY"}
        Dim seq2 = {"banana"}
        Dim res = seq1.Except(seq2, StringComparer.OrdinalIgnoreCase)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["APPLE,CHERRY"]);
}

#[test]
fn test_vb_linq_union_by_key_selector() {
    let src = r#"
Imports System.Linq

Class Person
    Public Property Name As String
    Public Sub New(n As String) : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim list1 = {New Person("Alice"), New Person("Bob")}
        Dim list2 = {New Person("Bob"), New Person("Charlie")}
        Dim res = list1.UnionBy(list2, Function(p) p.Name)
        For Each p In res
            Console.WriteLine(p.Name)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alice", "Bob", "Charlie"]);
}

#[test]
fn test_vb_linq_intersect_by_key_selector() {
    let src = r#"
Imports System.Linq

Class Item
    Public Property ID As Integer
    Public Sub New(id As Integer) : Me.ID = id : End Sub
End Class

Module Program
    Sub Main()
        Dim list1 = {New Item(1), New Item(2), New Item(3)}
        Dim keys2 = {2, 3, 4}
        Dim res = list1.IntersectBy(keys2, Function(i) i.ID)
        For Each i In res
            Console.WriteLine(i.ID)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "3"]);
}

#[test]
fn test_vb_linq_except_by_key_selector() {
    let src = r#"
Imports System.Linq

Class Product
    Public Property SKU As String
    Public Sub New(s As String) : SKU = s : End Sub
End Class

Module Program
    Sub Main()
        Dim prods = {New Product("A101"), New Product("B202"), New Product("C303")}
        Dim excludedSkus = {"B202"}
        Dim res = prods.ExceptBy(excludedSkus, Function(p) p.SKU)
        For Each p In res
            Console.WriteLine(p.SKU)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A101", "C303"]);
}

#[test]
fn test_vb_linq_union_empty_with_non_empty() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim empty As Integer() = {}
        Dim nonArr = {10, 20}
        Dim res = empty.Union(nonArr)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}

#[test]
fn test_vb_linq_intersect_no_common_elements() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2}
        Dim seq2 = {3, 4}
        Dim res = seq1.Intersect(seq2)
        Console.WriteLine(res.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_linq_except_disjoint_sets() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2}
        Dim seq2 = {3, 4}
        Dim res = seq1.Except(seq2)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2"]);
}

#[test]
fn test_vb_linq_union_identical_sets() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2, 3}
        Dim seq2 = {1, 2, 3}
        Dim res = seq1.Union(seq2)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_linq_intersect_identical_sets() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2, 3}
        Dim seq2 = {1, 2, 3}
        Dim res = seq1.Intersect(seq2)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_linq_except_identical_sets_empty_result() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2, 3}
        Dim seq2 = {1, 2, 3}
        Dim res = seq1.Except(seq2)
        Console.WriteLine(res.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_linq_union_tuples() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim set1 = {("A", 1), ("B", 2)}
        Dim set2 = {("B", 2), ("C", 3)}
        Dim res = set1.Union(set2)
        Console.WriteLine(res.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_linq_intersect_structs() {
    let src = r#"
Imports System.Linq

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer) : Me.X = x : Me.Y = y : End Sub
End Structure

Module Program
    Sub Main()
        Dim set1 = {New Point(0, 0), New Point(1, 1)}
        Dim set2 = {New Point(1, 1), New Point(2, 2)}
        Dim res = set1.Intersect(set2)
        Console.WriteLine(res.First().X & "," & res.First().Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,1"]);
}

#[test]
fn test_vb_linq_sequence_equal_operator() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2, 3}
        Dim seq2 = {1, 2, 3}
        Dim seq3 = {3, 2, 1}
        Console.WriteLine(seq1.SequenceEqual(seq2) & "|" & seq1.SequenceEqual(seq3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_linq_sequence_equal_custom_comparer() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {"a", "b"}
        Dim seq2 = {"A", "B"}
        Console.WriteLine(seq1.SequenceEqual(seq2, StringComparer.OrdinalIgnoreCase))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_linq_set_operations_pipeline_chain() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim setA = {1, 2, 3, 4}
        Dim setB = {3, 4, 5, 6}
        Dim setC = {4, 5}
        ' (A union B) except C -> {1,2,3,4,5,6} except {4,5} -> {1,2,3,6}
        Dim res = setA.Union(setB).Except(setC)
        Console.WriteLine(String.Join(",", res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,6"]);
}
