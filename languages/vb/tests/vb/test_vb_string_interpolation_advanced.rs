use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String Interpolation (Advanced)
// ═══════════════════════════════════════════════════════════

#[test]
fn string_interpolation_formatting() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim price As Double = 12.345
        ' Interpolation with formatting components like width and format string
        ' {expression,alignment:formatString}
        Dim result As String = $"{price,10:F2}"
        Console.WriteLine(result.Trim())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["12.35"]);
}

#[test]
fn string_interpolation_ternary() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim isValid As Boolean = True
        ' Interpolation containing expressions like If()
        Dim result As String = $"Status: {If(isValid, "Valid", "Invalid")}"
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Status: Valid"]);
}
