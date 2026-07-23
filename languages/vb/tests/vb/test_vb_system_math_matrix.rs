use super::helpers::run_vb;

#[test]
fn math_abs_is_non_negative() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Abs(-12))
        Console.WriteLine(Math.Abs(0))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["12", "0"]);
}

#[test]
fn math_min_and_max_contracts() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Min(11, 9))
        Console.WriteLine(Math.Max(11, 9))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["9", "11"]);
}

#[test]
fn math_floor_and_ceil_round() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Floor(3.9))
        Console.WriteLine(Math.Ceiling(3.1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn math_truncate_removes_fraction() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Truncate(9.9))
        Console.WriteLine(Math.Truncate(-9.9))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["9", "-9"]);
}

#[test]
fn math_sign_distinguishes_negative_zero_positive() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Sign(-4))
        Console.WriteLine(Math.Sign(0))
        Console.WriteLine(Math.Sign(19))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["-1", "0", "1"]);
}

#[test]
fn math_pow_and_round_is_reversible_for_small_input() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Pow(2, 3))
        Console.WriteLine(Math.Round(Math.Pow(2, 2) + 0.5))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["8", "4"]);
}

#[test]
fn math_sqrt_roundtrip_to_integer_square() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim root As Double = Math.Sqrt(64)
        Console.WriteLine(root)
        Console.WriteLine(root * root)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["8", "64"]);
}

#[test]
fn math_trig_basic_identities() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Round(Math.Sin(Math.PI / 2)))
        Console.WriteLine(Math.Round(Math.Cos(0)))
        Console.WriteLine(Math.Round(Math.Tan(Math.PI / 4)))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "1", "1"]);
}

#[test]
fn math_log10_of_thousand_is_three() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Log10(1000))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3"]);
}

#[test]
fn math_exp_and_log_inverse() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Round(Math.Log(Math.Exp(2)), 5))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn math_log2_of_power_of_two() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Log2(8))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3"]);
}

#[test]
fn math_atan2_returns_expected_angle() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Round(Math.Atan2(1, 1), 4))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0.7854"]);
}

#[test]
fn math_acos_and_asin_roundtrip() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Round(Math.Acos(1), 4))
        Console.WriteLine(Math.Round(Math.Asin(0), 4))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn math_hyperbolic_is_stable_for_zero() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Sinh(0))
        Console.WriteLine(Math.Cosh(0))
        Console.WriteLine(Math.Tanh(0))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0", "1", "0"]);
}

#[test]
fn math_clamp_contract() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Clamp(5, 1, 10))
        Console.WriteLine(Math.Clamp(-2, 1, 10))
        Console.WriteLine(Math.Clamp(12, 1, 10))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "1", "10"]);
}
