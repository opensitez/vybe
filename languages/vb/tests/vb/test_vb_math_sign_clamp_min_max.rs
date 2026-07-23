use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Math & MathF Utilities (Sign, Clamp, Min, Max)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_math_sign_negative_zero_positive() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Sign(-45) & "|" & Math.Sign(0) & "|" & Math.Sign(99))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1|0|1"]);
}

#[test]
fn test_vb_math_clamp_within_below_above_range() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Clamp(50, 0, 100) & "|" & Math.Clamp(-10, 0, 100) & "|" & Math.Clamp(150, 0, 100))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["50|0|100"]);
}

#[test]
fn test_vb_math_min_max_overloads_double() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Min(3.14, 2.71) & "|" & Math.Max(3.14, 2.71))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2.71|3.14"]);
}

#[test]
fn test_vb_math_min_max_overloads_integer() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Min(-10, 5) & "|" & Math.Max(-10, 5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-10|5"]);
}

#[test]
fn test_vb_math_abs_overloads() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Abs(-100) & "|" & Math.Abs(-3.14159))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100|3.14159"]);
}

#[test]
fn test_vb_math_ceiling_and_floor() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Ceiling(4.1) & "|" & Math.Floor(4.9) & "|" & Math.Floor(-2.1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5|4|-3"]);
}

#[test]
fn test_vb_math_truncate_drops_fraction() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Truncate(5.99) & "|" & Math.Truncate(-5.99))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5|-5"]);
}

#[test]
fn test_vb_math_round_bankers_rounding() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' Banker's Rounding (Round to Even): 2.5 -> 2, 3.5 -> 4
        Console.WriteLine(Math.Round(2.5) & "|" & Math.Round(3.5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|4"]);
}

#[test]
fn test_vb_math_round_midpoint_away_from_zero() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Round(2.5, MidpointRounding.AwayFromZero) & "|" & Math.Round(3.5, MidpointRounding.AwayFromZero))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3|4"]);
}

#[test]
fn test_vb_math_round_with_digits() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim pi = 3.14159265
        Console.WriteLine(Math.Round(pi, 2) & "|" & Math.Round(pi, 4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3.14|3.1416"]);
}

#[test]
fn test_vb_math_pow_and_sqrt() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Pow(2, 8) & "|" & Math.Sqrt(144))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["256|12"]);
}

#[test]
fn test_vb_math_log_log10_e() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Log10(100) & "|" & Math.Round(Math.Log(Math.E), 2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|1"]);
}

#[test]
fn test_vb_math_trigonometric_functions() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim sin0 = Math.Sin(0)
        Dim cos0 = Math.Cos(0)
        Console.WriteLine(sin0 & "|" & cos0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0|1"]);
}

#[test]
fn test_vb_math_atan2_polar_coordinates() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim angle = Math.Atan2(1.0, 1.0)
        Console.WriteLine(Math.Round(angle, 4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0.7854"]);
}

#[test]
fn test_vb_math_divrem_integer() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim remVal As Integer
        Dim quot = Math.DivRem(17, 5, remVal)
        Console.WriteLine(quot & " R " & remVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3 R 2"]);
}

#[test]
fn test_vb_mathf_single_precision_functions() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim sqrtVal As Single = MathF.Sqrt(16.0F)
        Console.WriteLine(sqrtVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}

#[test]
fn test_vb_math_constants_pi_and_e() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Math.Round(Math.PI, 2) & "|" & Math.Round(Math.E, 2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3.14|2.72"]);
}

#[test]
fn test_vb_math_exp_exponential() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim res = Math.Exp(1)
        Console.WriteLine(Math.Round(res, 4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2.7183"]);
}

#[test]
fn test_vb_math_scaleb_floating_point() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' ScaleB(x, n) calculates x * 2^n
        Dim res = Math.ScaleB(1.5, 3) ' 1.5 * 8 = 12
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12"]);
}

#[test]
fn test_vb_math_ilogb_exponent_extraction() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim exp = Math.ILogB(1024.0)
        Console.WriteLine(exp)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}
