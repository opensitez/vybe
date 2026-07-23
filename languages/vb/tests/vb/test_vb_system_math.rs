use super::helpers::run_vb;

#[test]
fn math_pow_raises_to_power() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Pow(2, 10))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1024"]);
}

#[test]
fn math_sqrt_extracts_square_root() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Sqrt(144))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["12"]);
}

#[test]
fn math_log_and_exp_are_inverse() {
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
fn math_sin_of_half_pi_is_one() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Round(Math.Sin(Math.PI / 2)))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn math_cos_of_zero_is_one() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Cos(0))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn math_tan_of_pi_over_four_approx_one() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Round(Math.Tan(Math.PI / 4)))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn math_truncate_removes_fractional_part() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Truncate(9.9))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["9"]);
}

#[test]
fn math_floor_and_ceiling_behave_differently() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Floor(3.8))
        Console.WriteLine(Math.Ceiling(3.2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn math_sign_reports_negative_zero_and_positive() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Sign(-12))
        Console.WriteLine(Math.Sign(0))
        Console.WriteLine(Math.Sign(9))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["-1", "0", "1"]);
}

#[test]
fn math_min_and_max_return_extrema() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.Min(3, 7))
        Console.WriteLine(Math.Max(3, 7))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "7"]);
}

#[test]
fn math_constants_are_reasonable() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(Math.PI > 3.14 AndAlso Math.PI < 3.15)
        Console.WriteLine(Math.E > 2.71 AndAlso Math.E < 2.72)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn math_atan2_returns_correct_quadrant_angle() {
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
