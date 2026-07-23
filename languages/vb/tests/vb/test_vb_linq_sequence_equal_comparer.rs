use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ SequenceEqual & Custom Equality Comparers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_sequence_equal_primitives() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2, 3}
        Dim seq2 = {1, 2, 3}
        Dim seq3 = {1, 2, 4}

        Console.WriteLine(seq1.SequenceEqual(seq2))
        Console.WriteLine(seq1.SequenceEqual(seq3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_linq_sequence_equal_string_comparer() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {"apple", "BANANA"}
        Dim seq2 = {"APPLE", "banana"}

        Console.WriteLine(seq1.SequenceEqual(seq2, StringComparer.OrdinalIgnoreCase))
        Console.WriteLine(seq1.SequenceEqual(seq2, StringComparer.Ordinal))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}
