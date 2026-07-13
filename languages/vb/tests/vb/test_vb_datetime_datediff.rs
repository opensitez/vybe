use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DateDiff Math
// ═══════════════════════════════════════════════════════════

#[test]
fn datetime_datediff_days() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d1 As Date = #1/1/2026#
        Dim d2 As Date = #1/10/2026#
        ' Difference in days (d2 - d1)
        Console.WriteLine(DateDiff(DateInterval.Day, d1, d2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn datetime_datediff_hours_negative() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d1 As Date = #1/2/2026 12:00:00 PM#
        Dim d2 As Date = #1/2/2026 8:00:00 AM#
        ' Difference in hours (d2 - d1), should be negative
        Console.WriteLine(DateDiff("h", d1, d2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["-4"]);
}
