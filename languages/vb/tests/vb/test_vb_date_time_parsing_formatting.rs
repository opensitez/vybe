use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DateTime Parsing & Custom Format Strings
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_date_time_parse_exact_format() {
    let src = r#"
Imports System
Imports System.Globalization

Module Program
    Sub Main()
        Dim strVal As String = "2025-05-15 14:30:00"
        Dim dt As DateTime = DateTime.ParseExact(strVal, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
        Console.WriteLine(dt.Year & "-" & dt.Month & "-" & dt.Day & " " & dt.Hour & ":" & dt.Minute)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-5-15 14:30"]);
}

#[test]
fn test_vb_date_time_try_parse_success_and_failure() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As DateTime
        Dim ok1 As Boolean = DateTime.TryParse("2025-01-01", dt)
        Dim ok2 As Boolean = DateTime.TryParse("invalid date", dt)
        Console.WriteLine(ok1)
        Console.WriteLine(ok2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}
