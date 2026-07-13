use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DateAdd Math
// ═══════════════════════════════════════════════════════════

#[test]
fn datetime_dateadd_days() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim dt As Date = #1/1/2026#
        ' Add 10 days
        Dim newDt As Date = DateAdd(DateInterval.Day, 10, dt)
        Console.WriteLine(newDt.Day)
        Console.WriteLine(newDt.Month)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["11", "1"]);
}

#[test]
fn datetime_dateadd_months_negative() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim dt As Date = #5/15/2026#
        ' Subtract 2 months (DateInterval.Month)
        Dim newDt As Date = DateAdd("m", -2, dt)
        Console.WriteLine(newDt.Month)
        Console.WriteLine(newDt.Year)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "2026"]);
}
