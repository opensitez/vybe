use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Math Builtins (Rounding, Int, Fix)
// ═══════════════════════════════════════════════════════════

#[test]
fn math_int_vs_fix() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' For positive numbers, Int and Fix are identical
        Console.WriteLine(Int(9.9))
        Console.WriteLine(Fix(9.9))
        
        ' For negative numbers, Int rounds down, Fix truncates
        Console.WriteLine(Int(-9.9))
        Console.WriteLine(Fix(-9.9))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["9", "9", "-10", "-9"]);
}

#[test]
fn math_abs_sign() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Abs(-50.5))
        
        ' Sign returns -1, 0, or 1
        Console.WriteLine(Sign(-100))
        Console.WriteLine(Sign(0))
        Console.WriteLine(Sign(45))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["50.5", "-1", "0", "1"]);
}
