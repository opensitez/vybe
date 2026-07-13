use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Collection Initializers (Dictionary)
// ═══════════════════════════════════════════════════════════

#[test]
fn collection_initializer_dictionary() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        ' Dictionary collection initializer syntax (uses nested braces)
        Dim dict As New Dictionary(Of String, Integer) From {
            {"A", 1},
            {"B", 2},
            {"C", 3}
        }
        
        Console.WriteLine(dict.Count)
        Console.WriteLine(dict("B"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "2"]);
}
