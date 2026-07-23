use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DateTimeOffset & UTC Conversions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_date_time_offset_utc_now_conversion() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto As New DateTimeOffset(2025, 5, 1, 12, 0, 0, TimeSpan.FromHours(2))
        Dim utcDto = dto.ToUniversalTime()
        Console.WriteLine(dto.Offset.TotalHours)
        Console.WriteLine(utcDto.Offset.TotalHours)
        Console.WriteLine(utcDto.Hour)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "0", "10"]);
}

#[test]
fn test_vb_date_time_offset_equality() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dto1 As New DateTimeOffset(2025, 1, 1, 12, 0, 0, TimeSpan.FromHours(0))
        Dim dto2 As New DateTimeOffset(2025, 1, 1, 14, 0, 0, TimeSpan.FromHours(2))
        Console.WriteLine(dto1 = dto2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
