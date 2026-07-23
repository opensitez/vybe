use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ Single, SingleOrDefault, First, FirstOrDefault
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_single_matching_element() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10}
        Console.WriteLine(numbers.Single())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_linq_single_with_predicate() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 3, 5, 8, 9}
        Dim even = numbers.Single(Function(n) n Mod 2 = 0)
        Console.WriteLine(even)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_linq_single_or_default_matching() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10}
        Console.WriteLine(numbers.SingleOrDefault())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_linq_single_or_default_empty_returns_default() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim empty As Integer() = {}
        Console.WriteLine(empty.SingleOrDefault())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_linq_single_or_default_custom_default_value() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim empty As Integer() = {}
        ' LINQ SingleOrDefault overload accepting default fallback value
        Console.WriteLine(empty.SingleOrDefault(-1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1"]);
}

#[test]
fn test_vb_linq_single_throws_invalid_operation_when_empty() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim empty As Integer() = {}
        Try
            Dim x = empty.Single()
        Catch ex As InvalidOperationException
            Console.WriteLine("Single Empty Exception Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Single Empty Exception Caught"]);
}

#[test]
fn test_vb_linq_single_throws_invalid_operation_when_multiple_matches() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20}
        Try
            Dim x = numbers.Single()
        Catch ex As InvalidOperationException
            Console.WriteLine("Single Multiple Matches Exception Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Single Multiple Matches Exception Caught"]
    );
}

#[test]
fn test_vb_linq_single_or_default_throws_when_multiple_matches() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20}
        Try
            Dim x = numbers.SingleOrDefault()
        Catch ex As InvalidOperationException
            Console.WriteLine("SingleOrDefault Multiple Matches Exception Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["SingleOrDefault Multiple Matches Exception Caught"]
    );
}

#[test]
fn test_vb_linq_first_or_default_returns_first_matching() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20, 30}
        Console.WriteLine(numbers.FirstOrDefault())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_linq_first_or_default_predicate() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "banana", "cherry"}
        Dim found = words.FirstOrDefault(Function(w) w.StartsWith("b"))
        Console.WriteLine(found)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["banana"]);
}

#[test]
fn test_vb_linq_first_or_default_custom_fallback() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "cherry"}
        Dim found = words.FirstOrDefault(Function(w) w.StartsWith("z"), "FallbackWord")
        Console.WriteLine(found)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FallbackWord"]);
}

#[test]
fn test_vb_linq_last_or_default_matching() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20, 30}
        Console.WriteLine(numbers.LastOrDefault())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30"]);
}

#[test]
fn test_vb_linq_last_or_default_predicate() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5, 6}
        Dim lastEven = numbers.LastOrDefault(Function(n) n Mod 2 = 0)
        Console.WriteLine(lastEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6"]);
}

#[test]
fn test_vb_linq_last_or_default_custom_fallback() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 3, 5}
        Dim lastEven = numbers.LastOrDefault(Function(n) n Mod 2 = 0, -99)
        Console.WriteLine(lastEven)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-99"]);
}

#[test]
fn test_vb_linq_element_at_index() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"A", "B", "C"}
        Console.WriteLine(words.ElementAt(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B"]);
}

#[test]
fn test_vb_linq_element_at_or_default_valid_and_invalid() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"A", "B", "C"}
        Console.WriteLine(words.ElementAtOrDefault(1) & "|" & (words.ElementAtOrDefault(10) Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B|True"]);
}

#[test]
fn test_vb_linq_element_at_or_default_custom_fallback() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20}
        Console.WriteLine(numbers.ElementAtOrDefault(5, -1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1"]);
}

#[test]
fn test_vb_linq_single_reference_type_null_check() {
    let src = r#"
Imports System.Linq

Class Document
    Public Title As String = "Doc"
End Class

Module Program
    Sub Main()
        Dim docs = {New Document()}
        Dim doc = docs.SingleOrDefault()
        Console.WriteLine(doc.Title)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Doc"]);
}

#[test]
fn test_vb_linq_single_or_default_struct_type() {
    let src = r#"
Imports System.Linq

Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Module Program
    Sub Main()
        Dim empty As Point() = {}
        Dim pt = empty.SingleOrDefault()
        Console.WriteLine(pt.X & "," & pt.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0,0"]);
}

#[test]
fn test_vb_linq_first_or_default_nullable_type() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim items As Nullable(Of Integer)() = {Nothing, 42, Nothing}
        Dim firstVal = items.FirstOrDefault(Function(i) i.HasValue)
        Console.WriteLine(firstVal.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}
