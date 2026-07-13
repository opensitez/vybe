use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: GetType Generics
// ═══════════════════════════════════════════════════════════

#[test]
fn gettype_generics() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        ' GetType for a constructed generic type
        Dim t1 As Type = GetType(List(Of String))
        Console.WriteLine(t1.Name)
        
        ' GetType for an open generic type uses (Of )
        Dim t2 As Type = GetType(List(Of ))
        Console.WriteLine(t2.Name)
        
        ' GetType for multi-parameter open generic type
        Dim t3 As Type = GetType(Dictionary(Of , ))
        Console.WriteLine(t3.Name)
    End Sub
End Module
"#,
    );
    // Names of open generics usually look like "List`1" and "Dictionary`2"
    assert_eq!(out, vec!["List`1", "List`1", "Dictionary`2"]);
}
