use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: TimeZoneInfo & UTC Conversions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_time_zone_info_utc_id() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim utcZone = TimeZoneInfo.Utc
        Console.WriteLine(utcZone.Id)
        Console.WriteLine(utcZone.BaseUtcOffset.TotalHours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["UTC", "0"]);
}

#[test]
fn test_vb_time_zone_info_convert_time_utc() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dtUtc As New DateTime(2025, 6, 1, 12, 0, 0, DateTimeKind.Utc)
        Dim localDt = TimeZoneInfo.ConvertTime(dtUtc, TimeZoneInfo.Utc)
        Console.WriteLine(localDt.Hour)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12"]);
}
