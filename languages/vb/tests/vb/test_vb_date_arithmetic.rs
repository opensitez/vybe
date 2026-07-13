use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Date/Time Arithmetic
// ═══════════════════════════════════════════════════════════

#[test]
fn date_arithmetic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d1 As Date = #1/1/2024#
        Dim d2 As Date = #1/5/2024#
        
        ' Subtracting dates returns a TimeSpan
        Dim ts As TimeSpan = d2 - d1
        Console.WriteLine(ts.Days)
        
        ' Adding TimeSpan to Date
        Dim d3 As Date = d1 + New TimeSpan(10, 0, 0, 0)
        Console.WriteLine(d3.Day)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["4", "11"]);
}
