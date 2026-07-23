use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: LINQ SelectMany / Compound From Queries
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_linq_compound_from_flatten_nested_collections() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim categories = {
            New With {.Name = "Fruit", .Items = {"Apple", "Banana"}},
            New With {.Name = "Veggie", .Items = {"Carrot", "Pea"}}
        }

        Dim allItems = From cat In categories
                       From item In cat.Items
                       Select item

        Console.WriteLine(String.Join(",", allItems))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Apple,Banana,Carrot,Pea"]);
}

#[test]
fn test_vb_linq_select_many_method_syntax() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"cat", "dog"}
        Dim chars = words.SelectMany(Function(w) w.ToCharArray())
        Console.WriteLine(String.Join(",", chars))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["c,a,t,d,o,g"]);
}

#[test]
fn test_vb_linq_select_many_with_index() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim sentences = {"Hello World", "VB NET"}
        Dim result = sentences.SelectMany(Function(s, idx) s.Split(" "c).Select(Function(w) idx & ":" & w))
        Console.WriteLine(String.Join(",", result))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0:Hello,0:World,1:VB,1:NET"]);
}
