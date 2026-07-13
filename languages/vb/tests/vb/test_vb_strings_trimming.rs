use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Strings Advanced (LTrim, RTrim, Trim)
// ═══════════════════════════════════════════════════════════

#[test]
fn string_trim_functions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "   padded   "
        
        Console.WriteLine("[" & LTrim(text) & "]")
        Console.WriteLine("[" & RTrim(text) & "]")
        Console.WriteLine("[" & Trim(text) & "]")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["[padded   ]", "[   padded]", "[padded]"]);
}
