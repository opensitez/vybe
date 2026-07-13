use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DatePart Math
// ═══════════════════════════════════════════════════════════

#[test]
fn datetime_datepart_extract() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim dt As Date = #12/25/2026 14:30:45#
        Console.WriteLine(DatePart(DateInterval.Year, dt))
        Console.WriteLine(DatePart(DateInterval.Month, dt))
        Console.WriteLine(DatePart(DateInterval.Day, dt))
        Console.WriteLine(DatePart("h", dt))
        Console.WriteLine(DatePart("n", dt)) ' Minutes
        Console.WriteLine(DatePart("s", dt)) ' Seconds
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2026", "12", "25", "14", "30", "45"]);
}

#[test]
fn datetime_datepart_weekday() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim dt As Date = #7/4/2026# ' A Saturday
        ' 7 = Saturday for vbSunday start
        Console.WriteLine(DatePart(DateInterval.Weekday, dt))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["7"]);
}
