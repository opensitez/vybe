use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Hex / Octal Literals Advanced
// ═══════════════════════════════════════════════════════════

#[test]
fn hex_octal_literals_advanced() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Hex literal
        Dim h As Integer = &HFF
        
        ' Octal literal
        Dim o As Integer = &O77
        
        ' Binary literal (VB 15 / VB.NET 2017+)
        Dim b As Integer = &B1010
        
        Console.WriteLine(h)
        Console.WriteLine(o)
        Console.WriteLine(b)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["255", "63", "10"]);
}
