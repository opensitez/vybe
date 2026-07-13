use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Decimal Literals
// ═══════════════════════════════════════════════════════════

#[test]
fn decimal_literals_d_suffix() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' The D suffix specifies a Decimal literal
        Dim dec As Decimal = 12.345D
        Dim bigDec = 9999999999999999999.99D
        
        Console.WriteLine(dec.GetType().Name)
        Console.WriteLine(bigDec.GetType().Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Decimal", "Decimal"]);
}

#[test]
fn decimal_literals_at_character() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' The @ type character specifies a Decimal literal
        Dim dec = 12.345@
        
        Console.WriteLine(dec.GetType().Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Decimal"]);
}
