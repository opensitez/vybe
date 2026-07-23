use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Double & Single Special Values (NaN, Infinity, Epsilon)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_double_nan_checks() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim nanVal = Double.NaN
        Console.WriteLine(Double.IsNaN(nanVal) & "|" & (nanVal = nanVal))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_double_positive_infinity() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim posInf = Double.PositiveInfinity
        Console.WriteLine(Double.IsPositiveInfinity(posInf) & "|" & Double.IsInfinity(posInf))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_double_negative_infinity() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim negInf = Double.NegativeInfinity
        Console.WriteLine(Double.IsNegativeInfinity(negInf) & "|" & Double.IsInfinity(negInf))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_divide_by_zero_double_returns_infinity() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Double = 1.0
        Dim b As Double = 0.0
        Dim res = a / b
        Console.WriteLine(Double.IsPositiveInfinity(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_zero_divided_by_zero_double_returns_nan() {
    let src = r#"
Module Program
    Sub Main()
        Dim a As Double = 0.0
        Dim b As Double = 0.0
        Dim res = a / b
        Console.WriteLine(Double.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_single_nan_and_infinity_properties() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim sNaN = Single.NaN
        Dim sInf = Single.PositiveInfinity
        Console.WriteLine(Single.IsNaN(sNaN) & "|" & Single.IsInfinity(sInf))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_double_epsilon_smallest_positive() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim eps = Double.Epsilon
        Console.WriteLine(eps > 0.0 & "|" & eps < 0.00001)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_double_min_max_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Double.MinValue < 0.0 & "|" & Double.MaxValue > 0.0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_double_is_finite_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim normalVal = 3.14159
        Dim infVal = Double.PositiveInfinity
        Dim nanVal = Double.NaN
        Console.WriteLine(Double.IsFinite(normalVal) & "|" & Double.IsFinite(infVal) & "|" & Double.IsFinite(nanVal))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False|False"]);
}

#[test]
fn test_vb_double_is_normal_subnormal_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim normalVal = 1.0
        Dim zeroVal = 0.0
        Console.WriteLine(Double.IsNormal(normalVal) & "|" & Double.IsNormal(zeroVal))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_infinity_arithmetic_operations() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim inf = Double.PositiveInfinity
        Dim resAdd = inf + 100
        Dim resMult = inf * 2
        Console.WriteLine(Double.IsPositiveInfinity(resAdd) & "|" & Double.IsPositiveInfinity(resMult))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_infinity_minus_infinity_yields_nan() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim inf1 = Double.PositiveInfinity
        Dim inf2 = Double.PositiveInfinity
        Dim res = inf1 - inf2
        Console.WriteLine(Double.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_infinity_multiplied_by_zero_yields_nan() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim inf = Double.PositiveInfinity
        Dim res = inf * 0.0
        Console.WriteLine(Double.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_double_try_parse_special_strings() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim val As Double
        Dim okNaN = Double.TryParse("NaN", val)
        Dim okInf = Double.TryParse("Infinity", val)
        Console.WriteLine(okNaN & "|" & okInf)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_double_to_string_special_values() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Double.NaN.ToString() & "|" & Double.PositiveInfinity.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NaN|Infinity"]);
}

#[test]
fn test_vb_sqrt_negative_number_yields_nan() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim res = Math.Sqrt(-1.0)
        Console.WriteLine(Double.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_log_negative_number_yields_nan() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim res = Math.Log(-5.0)
        Console.WriteLine(Double.IsNaN(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_log_zero_yields_negative_infinity() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim res = Math.Log(0.0)
        Console.WriteLine(Double.IsNegativeInfinity(res))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_double_negative_zero_detection() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim posZero = 0.0
        Dim negZero = -0.0
        Console.WriteLine((posZero = negZero) & "|" & (1.0 / posZero > 0) & "|" & (1.0 / negZero < 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_double_bit_converter_roundtrip() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim orig = 3.1415926535
        Dim bits = BitConverter.DoubleToInt64Bits(orig)
        Dim restored = BitConverter.Int64BitsToDouble(bits)
        Console.WriteLine(orig = restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
