use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ Aggregate with Seed & Accumulator
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_aggregate_simple_product() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5}
        Dim product = numbers.Aggregate(Function(acc, n) acc * n)
        Console.WriteLine(product)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["120"]);
}

#[test]
fn test_vb_linq_aggregate_with_seed() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "banana", "cherry"}
        Dim totalChars = words.Aggregate(0, Function(total, nextWord) total + nextWord.Length)
        Console.WriteLine(totalChars)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["17"]);
}

#[test]
fn test_vb_linq_aggregate_with_seed_and_result_selector() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4}
        Dim formatted = numbers.Aggregate(10, Function(acc, n) acc + n, Function(finalSum) "Total: " & finalSum)
        Console.WriteLine(formatted)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Total: 20"]);
}
