use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Operators (GetType)
// ═══════════════════════════════════════════════════════════

#[test]
fn operator_gettype() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' GetType returns a System.Type object for the specified type
        Dim t As Type = GetType(String)
        Console.WriteLine(t.Name)
        
        Dim t2 As Type = GetType(Integer)
        Console.WriteLine(t2.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["String", "Int32"]);
}
