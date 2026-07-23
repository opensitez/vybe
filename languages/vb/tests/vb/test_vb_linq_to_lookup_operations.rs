use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ ToLookup One-to-Many Indexing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_to_lookup_grouping() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "apricot", "banana", "blueberry", "cherry"}
        Dim lookup = words.ToLookup(Function(w) w(0))

        Console.WriteLine(lookup.Count)
        Console.WriteLine(String.Join(",", lookup("a"c)))
        Console.WriteLine(String.Join(",", lookup("b"c)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3", "apple,apricot", "banana,blueberry"]);
}

#[test]
fn test_vb_linq_to_lookup_element_selector() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim items = {
            New With {.Cat = "A", .Val = 10},
            New With {.Cat = "A", .Val = 20},
            New With {.Cat = "B", .Val = 30}
        }

        Dim lookup = items.ToLookup(Function(i) i.Cat, Function(i) i.Val)
        Console.WriteLine(String.Join(",", lookup("A")))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20"]);
}
