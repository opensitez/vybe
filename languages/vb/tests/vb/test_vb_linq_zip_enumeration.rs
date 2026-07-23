use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ Zip Method Enumeration
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_zip_equal_length_collections() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3}
        Dim words = {"One", "Two", "Three"}
        Dim zipped = numbers.Zip(words, Function(n, w) n & "=" & w)
        Console.WriteLine(String.Join(",", zipped))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1=One,2=Two,3=Three"]);
}

#[test]
fn test_vb_linq_zip_unequal_length_collections() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5}
        Dim letters = {"A", "B"}
        Dim zipped = numbers.Zip(letters, Function(n, l) n & l)
        Console.WriteLine(String.Join(",", zipped))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1A,2B"]);
}
