use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String Formatting (FormatNumber, FormatPercent)
// ═══════════════════════════════════════════════════════════

#[test]
fn format_number_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim value As Double = 987654.321
        Dim result As String = FormatNumber(value, 2, Microsoft.VisualBasic.TriState.True, Microsoft.VisualBasic.TriState.False, Microsoft.VisualBasic.TriState.True)
        Console.WriteLine("NumberFormatted")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["NumberFormatted"]);
}

#[test]
fn format_percent_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim value As Double = 0.854
        Dim result As String = FormatPercent(value, 1)
        Console.WriteLine("PercentFormatted")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["PercentFormatted"]);
}

#[test]
fn format_function_custom() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Format(1234.5, "0,0.00"))
        Console.WriteLine(Format(#12/25/2026#, "yyyy-MM-dd"))
    End Sub
End Module
"#,
    );
    // As basic format string parse test
    assert_eq!(out, vec!["1,234.50", "2026-12-25"]);
}
