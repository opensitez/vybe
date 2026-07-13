use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: String Formatting (FormatDateTime)
// ═══════════════════════════════════════════════════════════

#[test]
fn format_datetime_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim dt As Date = #1/1/2026 14:30:00#
        
        Dim f1 As String = FormatDateTime(dt, DateFormat.GeneralDate)
        Dim f2 As String = FormatDateTime(dt, DateFormat.LongDate)
        Dim f3 As String = FormatDateTime(dt, DateFormat.ShortDate)
        Dim f4 As String = FormatDateTime(dt, DateFormat.LongTime)
        Dim f5 As String = FormatDateTime(dt, DateFormat.ShortTime)
        
        Console.WriteLine("FormattedDates")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["FormattedDates"]);
}
