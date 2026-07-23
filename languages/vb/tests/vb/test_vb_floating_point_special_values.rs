use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Floating Point Special Values (NaN, Infinity, Epsilon)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_float_double_nan_property() {
    let src = r#"
Module Program
    Sub Main()
        Dim d As Double = Double.NaN
        Console.WriteLine(Double.IsNaN(d))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_double_positive_infinity() {
    let src = r#"
Module Program
    Sub Main()
        Dim posInf As Double = Double.PositiveInfinity
        Console.WriteLine(Double.IsPositiveInfinity(posInf))
        Console.WriteLine(Double.IsInfinity(posInf))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_float_double_negative_infinity() {
    let src = r#"
Module Program
    Sub Main()
        Dim negInf As Double = Double.NegativeInfinity
        Console.WriteLine(Double.IsNegativeInfinity(negInf))
        Console.WriteLine(Double.IsInfinity(negInf))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_float_division_by_zero_double() {
    let src = r#"
Module Program
    Sub Main()
        Dim res1 As Double = 1.0 / 0.0
        Dim res2 As Double = -1.0 / 0.0
        Dim res3 As Double = 0.0 / 0.0
        Console.WriteLine(Double.IsPositiveInfinity(res1))
        Console.WriteLine(Double.IsNegativeInfinity(res2))
        Console.WriteLine(Double.IsNaN(res3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True", "True"]);
}

#[test]
fn test_vb_float_single_nan_property() {
    let src = r#"
Module Program
    Sub Main()
        Dim s As Single = Single.NaN
        Console.WriteLine(Single.IsNaN(s))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_nan_comparison_not_equal() {
    let src = r#"
Module Program
    Sub Main()
        Dim nan1 As Double = Double.NaN
        Dim nan2 As Double = Double.NaN
        Console.WriteLine(nan1 = nan2)
        Console.WriteLine(nan1 <> nan2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "True"]);
}

#[test]
fn test_vb_float_double_epsilon() {
    let src = r#"
Module Program
    Sub Main()
        Dim eps As Double = Double.Epsilon
        Console.WriteLine(eps > 0.0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_single_epsilon() {
    let src = r#"
Module Program
    Sub Main()
        Dim eps As Single = Single.Epsilon
        Console.WriteLine(eps > 0.0F)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_infinity_arithmetic_pos_plus_num() {
    let src = r#"
Module Program
    Sub Main()
        Dim inf As Double = Double.PositiveInfinity
        Dim res As Double = inf + 100.0
        Console.WriteLine(Double.IsPositiveInfinity(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_infinity_arithmetic_pos_minus_pos() {
    let src = r#"
Module Program
    Sub Main()
        Dim inf1 As Double = Double.PositiveInfinity
        Dim inf2 As Double = Double.PositiveInfinity
        Dim res As Double = inf1 - inf2
        Console.WriteLine(Double.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_infinity_multiply_zero() {
    let src = r#"
Module Program
    Sub Main()
        Dim inf As Double = Double.PositiveInfinity
        Dim res As Double = inf * 0.0
        Console.WriteLine(Double.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_subnormal_number_underflow() {
    let src = r#"
Module Program
    Sub Main()
        Dim small As Double = 1e-320
        Dim halved As Double = small / 10.0
        Console.WriteLine(halved >= 0.0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_math_sqrt_negative() {
    let src = r#"
Module Program
    Sub Main()
        Dim res As Double = Math.Sqrt(-4.0)
        Console.WriteLine(Double.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_math_log_negative() {
    let src = r#"
Module Program
    Sub Main()
        Dim res As Double = Math.Log(-1.0)
        Console.WriteLine(Double.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_math_log_zero() {
    let src = r#"
Module Program
    Sub Main()
        Dim res As Double = Math.Log(0.0)
        Console.WriteLine(Double.IsNegativeInfinity(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_double_min_max_values() {
    let src = r#"
Module Program
    Sub Main()
        Dim minD As Double = Double.MinValue
        Dim maxD As Double = Double.MaxValue
        Console.WriteLine(minD < 0.0)
        Console.WriteLine(maxD > 0.0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_float_single_min_max_values() {
    let src = r#"
Module Program
    Sub Main()
        Dim minS As Single = Single.MinValue
        Dim maxS As Single = Single.MaxValue
        Console.WriteLine(minS < 0.0F)
        Console.WriteLine(maxS > 0.0F)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "True"]);
}

#[test]
fn test_vb_float_parse_infinity_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim inf As Double = Double.Parse("Infinity")
        Console.WriteLine(Double.IsPositiveInfinity(inf))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_parse_nan_string() {
    let src = r#"
Module Program
    Sub Main()
        Dim nanVal As Double = Double.Parse("NaN")
        Console.WriteLine(Double.IsNaN(nanVal))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_float_to_string_special_values() {
    let src = r#"
Module Program
    Sub Main()
        Console.WriteLine(Double.NaN.ToString())
        Console.WriteLine(Double.PositiveInfinity.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NaN", "Infinity"]);
}
