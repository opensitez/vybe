use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DateTime Properties (DayOfYear, Ticks, TimeOfDay)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_date_time_day_of_year_leap_year() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dtLeap As New DateTime(2024, 12, 31)
        Dim dtNorm As New DateTime(2025, 12, 31)
        Console.WriteLine(dtLeap.DayOfYear & "|" & dtNorm.DayOfYear)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["366|365"]);
}

#[test]
fn test_vb_date_time_ticks_precision() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc)
        Console.WriteLine(dt.Ticks > 0L)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_time_of_day_timespan() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 6, 15, 14, 30, 45)
        Dim tod = dt.TimeOfDay
        Console.WriteLine(tod.Hours & ":" & tod.Minutes & ":" & tod.Seconds)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["14:30:45"]);
}

#[test]
fn test_vb_date_time_day_of_week_enum() {
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

#[test]
fn test_vb_date_time_date_property_resets_time() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 5, 20, 18, 45, 12)
        Dim dOnly = dt.Date
        Console.WriteLine(dOnly.Hour & ":" & dOnly.Minute & ":" & dOnly.Second)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0:0:0"]);
}

#[test]
fn test_vb_date_time_add_days_and_hours() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 1, 10, 0, 0)
        Dim dtFuture = dt.AddDays(5).AddHours(3)
        Console.WriteLine(dtFuture.ToString("yyyy-MM-dd HH:mm"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-01-06 13:00"]);
}

#[test]
fn test_vb_date_time_add_months_year_rollover() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 11, 15)
        Dim dtNext = dt.AddMonths(3)
        Console.WriteLine(dtNext.Year & "-" & dtNext.Month)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2026-2"]);
}

#[test]
fn test_vb_date_time_subtract_date_time_yields_timespan() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 1, 10)
        Dim d2 As New DateTime(2025, 1, 1)
        Dim diff As TimeSpan = d1 - d2
        Console.WriteLine(diff.TotalDays)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9"]);
}

#[test]
fn test_vb_date_time_subtract_timespan() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 10)
        Dim dtPrev = dt.Subtract(TimeSpan.FromDays(2))
        Console.WriteLine(dtPrev.Day)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_date_time_constructor_ticks_overload() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim orig As New DateTime(2025, 7, 4, 12, 0, 0)
        Dim ticks = orig.Ticks
        Dim reconstructed As New DateTime(ticks)
        Console.WriteLine(orig = reconstructed)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_kind_utc_local_unspecified() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dtUtc As New DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc)
        Console.WriteLine(dtUtc.Kind.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Utc"]);
}

#[test]
fn test_vb_date_time_to_universal_time() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dtUtc As New DateTime(2025, 1, 1, 12, 0, 0, DateTimeKind.Utc)
        Dim converted = dtUtc.ToUniversalTime()
        Console.WriteLine(converted.Hour)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12"]);
}

#[test]
fn test_vb_date_time_to_short_date_string() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 4, 15)
        Console.WriteLine(dt.ToString("yyyy/MM/dd"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025/04/15"]);
}

#[test]
fn test_vb_date_time_days_in_month() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim feb2024 = DateTime.DaysInMonth(2024, 2)
        Dim feb2025 = DateTime.DaysInMonth(2025, 2)
        Console.WriteLine(feb2024 & "|" & feb2025)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["29|28"]);
}

#[test]
fn test_vb_date_time_min_max_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(DateTime.MinValue.Year & "|" & DateTime.MaxValue.Year)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|9999"]);
}

#[test]
fn test_vb_date_time_millisecond_precision() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 1, 10, 20, 30, 456)
        Console.WriteLine(dt.Millisecond)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["456"]);
}

#[test]
fn test_vb_date_time_equality_operators() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 1, 1)
        Dim d2 As New DateTime(2025, 1, 1)
        Dim d3 As New DateTime(2025, 1, 2)
        Console.WriteLine((d1 = d2) & "|" & (d1 < d3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_date_time_add_ticks() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 1)
        Dim dtPlusOneSec = dt.AddTicks(TimeSpan.TicksPerSecond)
        Console.WriteLine(dtPlusOneSec.Second)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_date_literal_syntax() {
    let src = r#"
Module Program
    Sub Main()
        Dim d As Date = #1/1/2025#
        Console.WriteLine(d.Year & "-" & d.Month & "-" & d.Day)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-1-1"]);
}

#[test]
fn test_vb_date_time_specify_kind() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 1)
        Dim dtUtc = DateTime.SpecifyKind(dt, DateTimeKind.Utc)
        Console.WriteLine(dtUtc.Kind.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Utc"]);
}
