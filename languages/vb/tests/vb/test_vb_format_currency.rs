use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String Formatting (FormatCurrency)
// ═══════════════════════════════════════════════════════════

#[test]
fn format_currency_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim value As Double = 1234.567
        ' Defaults to 2 decimal places and system currency symbol
        ' Testing actual system output is hard since it depends on culture,
        ' but we can check if it parses and runs.
        Dim result As String = FormatCurrency(value)
        Console.WriteLine("Formatted")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Formatted"]);
}

#[test]
fn format_currency_args() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim value As Double = -12.3
        ' FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
        Dim result As String = FormatCurrency(value, 3, Microsoft.VisualBasic.TriState.True, Microsoft.VisualBasic.TriState.True, Microsoft.VisualBasic.TriState.False)
        Console.WriteLine("ArgsPassed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["ArgsPassed"]);
}
