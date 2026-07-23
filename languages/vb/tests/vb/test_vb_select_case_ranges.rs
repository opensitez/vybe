use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Select Case Ranges
// ═══════════════════════════════════════════════════════════

#[test]
fn select_case_ranges_advanced() {
    let out = run_vb(
        r#"
Module M
    Sub ClassifyNumber(n As Integer)
        Select Case n
            Case 1 To 5
                Console.WriteLine("Small")
            Case 6 To 10, 15 To 20
                Console.WriteLine("Medium or Teen")
            Case Is > 100
                Console.WriteLine("Large")
            Case Else
                Console.WriteLine("Other")
        End Select
    End Sub

    Sub Main()
        ClassifyNumber(3)
        ClassifyNumber(8)
        ClassifyNumber(17)
        ClassifyNumber(150)
        ClassifyNumber(50)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec![
            "Small",
            "Medium or Teen",
            "Medium or Teen",
            "Large",
            "Other"
        ]
    );
}
