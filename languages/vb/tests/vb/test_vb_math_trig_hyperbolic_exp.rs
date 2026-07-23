use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Math Trigonometric, Hyperbolic & Exp
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_math_trig_functions() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Sin(0.0))
        Console.WriteLine(Math.Cos(0.0))
        Console.WriteLine(Math.Log10(100.0))
        Console.WriteLine(Math.Exp(0.0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0", "1", "2", "1"]);
}
