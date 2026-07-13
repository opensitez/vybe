use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Date/Time Literals (Advanced)
// ═══════════════════════════════════════════════════════════

#[test]
fn datetime_literals_advanced() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' ISO format YYYY-MM-DD
        Dim d1 As Date = #2024-05-15#
        
        ' With time YYYY-MM-DD HH:MM:SS
        Dim d2 As Date = #2024-05-15 14:30:00#
        
        ' AM/PM format
        Dim d3 As Date = #5/15/2024 2:30 PM#
        
        Console.WriteLine(d1.Year)
        Console.WriteLine(d2.Hour)
        Console.WriteLine(d3.Hour)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2024", "14", "14"]);
}
