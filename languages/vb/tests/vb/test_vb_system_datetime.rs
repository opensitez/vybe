use super::helpers::run_vb;

#[test]
fn system_datetime_parse_adddays_and_compare() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim parsed As DateTime = DateTime.Parse("2024-02-28T00:00:00")
        Dim leap As DateTime = parsed.AddDays(1)
        Console.WriteLine(DateTime.Compare(leap, parsed))
        Console.WriteLine(leap.Day)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "29"]);
}

#[test]
fn system_datetime_static_days_in_month_and_leap_year() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(DateTime.DaysInMonth(2000, 2))
        Console.WriteLine(DateTime.IsLeapYear(1999))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["29", "False"]);
}
