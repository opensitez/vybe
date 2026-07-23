use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DateTimeOffset & TimeZone UTC Offset Conversions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_date_time_offset_construction_and_offset() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim offset = TimeSpan.FromHours(-8)
        Dim dto As New DateTimeOffset(2025, 1, 1, 12, 0, 0, offset)
        Console.WriteLine(dto.Offset.Hours & "|" & dto.Hour)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-8|12"]);
}

#[test]
fn test_vb_date_time_offset_to_universal_time() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim offset = TimeSpan.FromHours(-5)
        Dim dto As New DateTimeOffset(2025, 1, 1, 12, 0, 0, offset)
        Dim utcDto = dto.ToUniversalTime()
        Console.WriteLine(utcDto.Hour & "|" & utcDto.Offset.Hours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["17|0"]);
}

#[test]
fn test_vb_date_time_offset_to_offset_conversion() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dtoPst As New DateTimeOffset(2025, 1, 1, 10, 0, 0, TimeSpan.FromHours(-8))
        Dim dtoEst = dtoPst.ToOffset(TimeSpan.FromHours(-5))
        Console.WriteLine(dtoEst.Hour & ":" & dtoEst.Minute)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["13:0"]);
}

#[test]
fn test_vb_date_time_offset_now_utcnow() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim utcNow = DateTimeOffset.UtcNow
        Console.WriteLine(utcNow.Offset.Ticks = 0L)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_offset_date_and_date_time_properties() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto As New DateTimeOffset(2025, 3, 15, 14, 30, 0, TimeSpan.FromHours(2))
        Console.WriteLine(dto.Date.ToString("yyyy-MM-dd") & "|" & dto.DateTime.Hour)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-03-15|14"]);
}

#[test]
fn test_vb_date_time_offset_utc_date_time_property() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto As New DateTimeOffset(2025, 1, 1, 10, 0, 0, TimeSpan.FromHours(-5))
        Dim utcDt = dto.UtcDateTime
        Console.WriteLine(utcDt.Hour & "|" & utcDt.Kind.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15|Utc"]);
}

#[test]
fn test_vb_date_time_offset_add_hours_preserves_offset() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto As New DateTimeOffset(2025, 1, 1, 10, 0, 0, TimeSpan.FromHours(3))
        Dim updated = dto.AddHours(5)
        Console.WriteLine(updated.Hour & "|" & updated.Offset.Hours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15|3"]);
}

#[test]
fn test_vb_date_time_offset_subtraction_yields_timespan() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto1 As New DateTimeOffset(2025, 1, 1, 18, 0, 0, TimeSpan.FromHours(0))
        Dim dto2 As New DateTimeOffset(2025, 1, 1, 10, 0, 0, TimeSpan.FromHours(-5)) ' Equivalent to 15:00 UTC
        Dim diff As TimeSpan = dto1 - dto2
        Console.WriteLine(diff.TotalHours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_date_time_offset_equality_same_utc_instant() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto1 As New DateTimeOffset(2025, 1, 1, 12, 0, 0, TimeSpan.FromHours(0))
        Dim dto2 As New DateTimeOffset(2025, 1, 1, 7, 0, 0, TimeSpan.FromHours(-5))
        Console.WriteLine(dto1 = dto2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_offset_equals_exact_method() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto1 As New DateTimeOffset(2025, 1, 1, 12, 0, 0, TimeSpan.FromHours(0))
        Dim dto2 As New DateTimeOffset(2025, 1, 1, 7, 0, 0, TimeSpan.FromHours(-5))
        ' EqualsExact checks both UTC instant AND Offset equality!
        Console.WriteLine(dto1.EqualsExact(dto2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_date_time_offset_unix_time_seconds_conversion() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim epoch = DateTimeOffset.FromUnixTimeSeconds(0L)
        Console.WriteLine(epoch.Year & "-" & epoch.Month & "-" & epoch.Day & " UTC")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1970-1-1 UTC"]);
}

#[test]
fn test_vb_date_time_offset_to_unix_time_milliseconds() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim epoch = DateTimeOffset.FromUnixTimeMilliseconds(1000L)
        Console.WriteLine(epoch.ToUnixTimeMilliseconds())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1000"]);
}

#[test]
fn test_vb_date_time_offset_parse_iso8601() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto = DateTimeOffset.Parse("2025-07-04T12:00:00+02:00")
        Console.WriteLine(dto.Year & "|" & dto.Offset.Hours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025|2"]);
}

#[test]
fn test_vb_date_time_offset_try_parse_iso8601() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto As DateTimeOffset
        Dim success = DateTimeOffset.TryParse("2025-12-31T23:59:59Z", dto)
        Console.WriteLine(success & "|" & dto.Offset.Hours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|0"]);
}

#[test]
fn test_vb_date_time_offset_comparison_operators() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim t1 = New DateTimeOffset(2025, 1, 1, 10, 0, 0, TimeSpan.Zero)
        Dim t2 = New DateTimeOffset(2025, 1, 1, 12, 0, 0, TimeSpan.Zero)
        Console.WriteLine((t1 < t2) & "|" & (t2 > t1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_date_time_offset_to_string_format_z() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto As New DateTimeOffset(2025, 5, 10, 15, 30, 0, TimeSpan.FromHours(5))
        Console.WriteLine(dto.ToString("yyyy-MM-dd HH:mm zzz"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-05-10 15:30 +05:00"]);
}

#[test]
fn test_vb_date_time_offset_min_max_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(DateTimeOffset.MinValue.Year & "|" & DateTimeOffset.MaxValue.Year)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|9999"]);
}

#[test]
fn test_vb_date_time_offset_ticks_property() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto As New DateTimeOffset(2025, 1, 1, 0, 0, 0, TimeSpan.Zero)
        Console.WriteLine(dto.Ticks > 0L)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_offset_implicit_conversion_from_datetime() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 6, 1, 10, 0, 0, DateTimeKind.Utc)
        Dim dto As DateTimeOffset = dt
        Console.WriteLine(dto.Hour & "|" & dto.Offset.Hours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10|0"]);
}

#[test]
fn test_vb_date_time_offset_hash_code_consistency() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto1 As New DateTimeOffset(2025, 1, 1, 12, 0, 0, TimeSpan.Zero)
        Dim dto2 As New DateTimeOffset(2025, 1, 1, 12, 0, 0, TimeSpan.Zero)
        Console.WriteLine(dto1.GetHashCode() = dto2.GetHashCode())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
