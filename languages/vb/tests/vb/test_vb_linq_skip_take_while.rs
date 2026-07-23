use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ SkipWhile, TakeWhile & Pagination Operators
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_takewhile_predicate() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5, 1, 2}
        Dim result = numbers.TakeWhile(Function(n) n < 4)
        Console.WriteLine(String.Join(",", result))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_linq_skipwhile_predicate() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5, 1, 2}
        Dim result = numbers.SkipWhile(Function(n) n < 4)
        Console.WriteLine(String.Join(",", result))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4,5,1,2"]);
}

#[test]
fn test_vb_linq_takewhile_indexed_predicate() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20, 30, 40, 50}
        ' Take elements while element >= index * 10
        Dim result = numbers.TakeWhile(Function(n, idx) n >= idx * 10)
        Console.WriteLine(String.Join(",", result))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30,40,50"]);
}

#[test]
fn test_vb_linq_skipwhile_indexed_predicate() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {5, 10, 15, 20, 25}
        ' Skip elements while value <= index * 10 (5<=0 False, so skips nothing)
        Dim result = numbers.SkipWhile(Function(n, idx) n <= idx * 10)
        Console.WriteLine(String.Join(",", result))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5,10,15,20,25"]);
}

#[test]
fn test_vb_linq_skip_and_take_pagination() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}
        Dim page2 = numbers.Skip(3).Take(3)
        Console.WriteLine(String.Join(",", page2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4,5,6"]);
}

#[test]
fn test_vb_linq_query_syntax_skip_take() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5, 6}
        Dim query = From n In numbers Take 3
        Console.WriteLine(String.Join(",", query))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_linq_query_syntax_skip_clause() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5, 6}
        Dim query = From n In numbers Skip 4
        Console.WriteLine(String.Join(",", query))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5,6"]);
}

#[test]
fn test_vb_linq_query_syntax_skip_while_clause() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "apricot", "banana", "cherry"}
        Dim query = From w In words Skip While w.StartsWith("a")
        Console.WriteLine(String.Join(",", query))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["banana,cherry"]);
}

#[test]
fn test_vb_linq_query_syntax_take_while_clause() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "apricot", "banana", "cherry"}
        Dim query = From w In words Take While w.StartsWith("a")
        Console.WriteLine(String.Join(",", query))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["apple,apricot"]);
}

#[test]
fn test_vb_linq_takewhile_all_match() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim nums = {2, 4, 6}
        Dim result = nums.TakeWhile(Function(n) n Mod 2 = 0)
        Console.WriteLine(String.Join(",", result))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2,4,6"]);
}

#[test]
fn test_vb_linq_takewhile_none_match() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim nums = {1, 3, 5}
        Dim result = nums.TakeWhile(Function(n) n Mod 2 = 0)
        Console.WriteLine(result.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_linq_skipwhile_all_match() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim nums = {2, 4, 6}
        Dim result = nums.SkipWhile(Function(n) n Mod 2 = 0)
        Console.WriteLine(result.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_linq_skip_exceeds_count() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim nums = {1, 2, 3}
        Dim result = nums.Skip(100)
        Console.WriteLine(result.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_linq_take_exceeds_count() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim nums = {1, 2, 3}
        Dim result = nums.Take(100)
        Console.WriteLine(String.Join(",", result))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3"]);
}

#[test]
fn test_vb_linq_takewhile_complex_objects() {
    let src = r#"
Imports System.Linq

Class Item
    Public Property Price As Double
    Public Sub New(p As Double) : Price = p : End Sub
End Class

Module Program
    Sub Main()
        Dim items = {New Item(10), New Item(20), New Item(50), New Item(5)}
        Dim cheap = items.TakeWhile(Function(i) i.Price < 30)
        Console.WriteLine(cheap.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_linq_skipwhile_complex_objects() {
    let src = r#"
Imports System.Linq

Class LogEntry
    Public Property Level As String
    Public Sub New(l As String) : Level = l : End Sub
End Class

Module Program
    Sub Main()
        Dim logs = {New LogEntry("INFO"), New LogEntry("INFO"), New LogEntry("ERROR"), New LogEntry("INFO")}
        Dim startingFromError = logs.SkipWhile(Function(l) l.Level = "INFO")
        Console.WriteLine(startingFromError.First().Level & "|" & startingFromError.Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ERROR|2"]);
}

#[test]
fn test_vb_linq_chunk_batches() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5}
        Dim chunks = numbers.Chunk(2)
        For Each chunk In chunks
            Console.WriteLine(String.Join("-", chunk))
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1-2", "3-4", "5"]);
}

#[test]
fn test_vb_linq_take_last_elements() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20, 30, 40, 50}
        Dim lastTwo = numbers.TakeLast(2)
        Console.WriteLine(String.Join(",", lastTwo))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["40,50"]);
}

#[test]
fn test_vb_linq_skip_last_elements() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {10, 20, 30, 40, 50}
        Dim allButLastTwo = numbers.SkipLast(2)
        Console.WriteLine(String.Join(",", allButLastTwo))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_linq_combination_skipwhile_takewhile() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim sequence = {1, 1, 2, 3, 5, 8, 13, 21}
        ' Skip ones, then take values under 10
        Dim subSeq = sequence.SkipWhile(Function(n) n = 1).TakeWhile(Function(n) n < 10)
        Console.WriteLine(String.Join(",", subSeq))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2,3,5,8"]);
}
