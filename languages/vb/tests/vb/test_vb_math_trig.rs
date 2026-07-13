use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Math Builtins (Trigonometry)
// ═══════════════════════════════════════════════════════════

#[test]
fn math_trigonometry_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Sin(0) = 0
        Console.WriteLine(Sin(0))
        ' Cos(0) = 1
        Console.WriteLine(Cos(0))
        ' Tan(0) = 0
        Console.WriteLine(Tan(0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "1", "0"]);
}

#[test]
fn math_trigonometry_atn() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Atn(0) = 0
        Console.WriteLine(Atn(0))
        ' Approximating PI / 4
        Console.WriteLine(Math.Round(Atn(1) * 4, 2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "3.14"]);
}
