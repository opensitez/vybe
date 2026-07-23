use super::helpers::run_vb;

#[test]
fn datetime_to_string_yyyy_mm_dd() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d As Date = New Date(2024, 6, 15)
        Console.WriteLine(d.ToString("yyyy-MM-dd"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2024-06-15"]);
}

#[test]
fn datetime_to_string_dd_slash_mm_slash_yyyy() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d As Date = New Date(2024, 1, 5)
        Console.WriteLine(d.ToString("dd/MM/yyyy"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["05/01/2024"]);
}

#[test]
fn datetime_to_string_hh_mm_ss() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d As New DateTime(2024, 1, 1, 13, 5, 9)
        Console.WriteLine(d.ToString("HH:mm:ss"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["13:05:09"]);
}

#[test]
fn datetime_parse_exact_with_invariant_culture() {
    let out = run_vb(
        r#"
Imports System.Globalization

Module M
    Sub Main()
        Dim d As Date = Date.ParseExact("2024-03-21", "yyyy-MM-dd", CultureInfo.InvariantCulture)
        Console.WriteLine(d.Year)
        Console.WriteLine(d.Month)
        Console.WriteLine(d.Day)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2024", "3", "21"]);
}

#[test]
fn datetime_try_parse_returns_false_for_invalid() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d As Date
        Console.WriteLine(Date.TryParse("not-a-date", d))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False"]);
}

#[test]
fn datetime_adddays_and_addhours_affect_calendar_values() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d As Date = New Date(2024, 2, 27)
        Dim next As Date = d.AddDays(1)
        Console.WriteLine(next.Day)
        Console.WriteLine(next.Month)

        Dim adjusted As Date = d.AddHours(3.5)
        Console.WriteLine(adjusted.Hour)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["28", "2", "3"]);
}

#[test]
fn datetime_today_is_not_min_value() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Date.Today <> Date.MinValue)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn datetime_time_of_day_from_ticks() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d As Date = New Date(2024, 1, 1, 5, 45, 10)
        Console.WriteLine(d.TimeOfDay.Hours)
        Console.WriteLine(d.TimeOfDay.Minutes)
        Console.WriteLine(d.TimeOfDay.Seconds)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "45", "10"]);
}

#[test]
fn datetime_compare_orders_by_calendar_value() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim a As Date = New Date(2024, 1, 1)
        Dim b As Date = New Date(2024, 1, 2)
        Console.WriteLine(a < b)
        Console.WriteLine(a.CompareTo(b) < 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn datetime_leap_year_days_in_month() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Date.DaysInMonth(2000, 2))
        Console.WriteLine(Date.IsLeapYear(1999))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["29", "False"]);
}
