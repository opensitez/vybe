use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: TimeSpan Arithmetic, Properties & Factory Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_time_span_total_units() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts As TimeSpan = TimeSpan.FromHours(2.5)
        Console.WriteLine(ts.TotalMinutes)
        Console.WriteLine(ts.TotalSeconds)
        Console.WriteLine(ts.Hours & ":" & ts.Minutes)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["150", "9000", "2:30"]);
}

#[test]
fn test_vb_time_span_date_subtraction() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 1, 10)
        Dim d2 As New DateTime(2025, 1, 1)
        Dim diff As TimeSpan = d1 - d2
        Console.WriteLine(diff.Days)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9"]);
}
