use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Date Math (AddDays, AddMonths, AddYears, IsLeapYear)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_date_time_add_methods() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2024, 2, 28) ' Leap year
        Console.WriteLine(dt.AddDays(1).ToString("yyyy-MM-dd"))
        Console.WriteLine(dt.AddMonths(1).ToString("yyyy-MM-dd"))
        Console.WriteLine(dt.AddYears(1).ToString("yyyy-MM-dd"))
        Console.WriteLine(DateTime.IsLeapYear(2024))
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["2024-02-29", "2024-03-28", "2025-02-28", "True"]
    );
}

#[test]
fn test_vb_date_time_day_of_week() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 1) ' Wednesday
        Console.WriteLine(dt.DayOfWeek.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Wednesday"]);
}
