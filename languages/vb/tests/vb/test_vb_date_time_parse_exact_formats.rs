use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DateTime.ParseExact & Custom Format Strings
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_date_time_parse_exact_single_format() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("2025-08-15", "yyyy-MM-dd", CultureInfo.InvariantCulture)
        Console.WriteLine(dt.Year & "-" & dt.Month & "-" & dt.Day)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-8-15"]);
}

#[test]
fn test_vb_date_time_parse_exact_multiple_formats_array() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim formats As String() = {"yyyy/MM/dd", "dd-MM-yyyy", "yyyy-MM-dd"}
        Dim dt = DateTime.ParseExact("15-08-2025", formats, CultureInfo.InvariantCulture, DateTimeStyles.None)
        Console.WriteLine(dt.Day & "/" & dt.Month & "/" & dt.Year)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15/8/2025"]);
}

#[test]
fn test_vb_date_time_try_parse_exact_success() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt As DateTime
        Dim ok = DateTime.TryParseExact("20251231", "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, dt)
        Console.WriteLine(ok & "|" & dt.Year & "-" & dt.Month & "-" & dt.Day)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|2025-12-31"]);
}

#[test]
fn test_vb_date_time_try_parse_exact_invalid_format_fails_gracefully() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt As DateTime
        Dim ok = DateTime.TryParseExact("InvalidDate", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, dt)
        Console.WriteLine(ok)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_date_time_parse_exact_time_format_with_am_pm() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("02:30 PM", "hh:mm tt", CultureInfo.InvariantCulture)
        Console.WriteLine(dt.Hour & ":" & dt.Minute)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["14:30"]);
}

#[test]
fn test_vb_date_time_parse_exact_24_hour_time() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("23:45:10", "HH:mm:ss", CultureInfo.InvariantCulture)
        Console.WriteLine(dt.Hour & ":" & dt.Minute & ":" & dt.Second)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["23:45:10"]);
}

#[test]
fn test_vb_date_time_parse_exact_with_milliseconds() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("12:00:00.789", "HH:mm:ss.fff", CultureInfo.InvariantCulture)
        Console.WriteLine(dt.Millisecond)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["789"]);
}

#[test]
fn test_vb_date_time_parse_exact_adjust_to_universal_style() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("2025-01-01 10:00:00Z", "yyyy-MM-dd HH:mm:ss'Z'", CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal)
        Console.WriteLine(dt.Kind.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Utc"]);
}

#[test]
fn test_vb_date_time_parse_exact_escaped_literals() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("Year: 2025, Month: 05", "'Year:' yyyy', Month:' MM", CultureInfo.InvariantCulture)
        Console.WriteLine(dt.Year & "/" & dt.Month)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025/5"]);
}

#[test]
fn test_vb_date_time_parse_exact_rfc1123_format() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("Wed, 01 Jan 2025 12:00:00 GMT", "r", CultureInfo.InvariantCulture)
        Console.WriteLine(dt.Year & "-" & dt.Month & "-" & dt.Day)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-1-1"]);
}

#[test]
fn test_vb_date_time_parse_exact_iso8601_o_format() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim str = "2025-06-15T14:30:00.0000000Z"
        Dim dt = DateTime.ParseExact(str, "o", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind)
        Console.WriteLine(dt.Year & "|" & dt.Kind.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025|Utc"]);
}

#[test]
fn test_vb_date_time_parse_exact_throws_format_exception() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Try
            DateTime.ParseExact("2025-13-40", "yyyy-MM-dd", CultureInfo.InvariantCulture)
        Catch ex As FormatException
            Console.WriteLine("FormatException Caught on Invalid Month/Day")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["FormatException Caught on Invalid Month/Day"]
    );
}

#[test]
fn test_vb_date_time_try_parse_custom_culture_german() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim culture = CultureInfo.GetCultureInfo("de-DE")
        Dim dt = DateTime.Parse("15.08.2025", culture)
        Console.WriteLine(dt.Day & "." & dt.Month & "." & dt.Year)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15.8.2025"]);
}

#[test]
fn test_vb_date_time_parse_exact_abbreviated_month_names() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("15-Aug-2025", "dd-MMM-yyyy", CultureInfo.InvariantCulture)
        Console.WriteLine(dt.Month)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_date_time_parse_exact_full_month_name() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("August 15, 2025", "MMMM dd, yyyy", CultureInfo.InvariantCulture)
        Console.WriteLine(dt.Month & "/" & dt.Day)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8/15"]);
}

#[test]
fn test_vb_date_time_parse_exact_assume_universal_style() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("2025-01-01 00:00:00", "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal)
        Console.WriteLine(dt.Kind.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Local"]);
}

#[test]
fn test_vb_date_time_parse_exact_short_year_format() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("25-01-01", "yy-MM-dd", CultureInfo.InvariantCulture)
        Console.WriteLine(dt.Year)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025"]);
}

#[test]
fn test_vb_date_time_parse_exact_white_space_handling() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("   2025-05-10   ", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces)
        Console.WriteLine(dt.Month & "-" & dt.Day)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5-10"]);
}

#[test]
fn test_vb_date_time_to_string_all_standard_formats() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 1, 15, 30, 0)
        Console.WriteLine(dt.ToString("s", CultureInfo.InvariantCulture))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-01-01T15:30:00"]);
}

#[test]
fn test_vb_date_time_parse_exact_unix_timestamp_string() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim dt = DateTime.ParseExact("2025-01-01T00:00:00.000Z", "yyyy-MM-ddTHH:mm:ss.fff'Z'", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal)
        Console.WriteLine(dt.Year & "|" & dt.Kind.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025|Utc"]);
}
