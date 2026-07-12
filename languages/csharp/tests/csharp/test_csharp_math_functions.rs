//! `System.Math` functions beyond basic rounding.
use super::helpers::run_csharp;

#[test]
fn math_pow_raises_base_to_integer_exponent() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.Pow(2, 10));"#),
        &["1024"]
    );
}

#[test]
fn math_sqrt_extracts_principal_square_root() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.Sqrt(144));"#),
        &["12"]
    );
}

#[test]
fn math_log_natural_is_inverse_of_exp() {
    assert_eq!(
        run_csharp(
            r#"double v = System.Math.Log(System.Math.E);
Console.WriteLine(System.Math.Round(v, 5));"#
        ),
        &["1"]
    );
}

#[test]
fn math_log10_of_thousand_equals_three() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.Log10(1000));"#),
        &["3"]
    );
}

#[test]
fn math_sin_of_half_pi_is_one() {
    assert_eq!(
        run_csharp(
            r#"double v = System.Math.Sin(System.Math.PI / 2);
Console.WriteLine(System.Math.Round(v));"#
        ),
        &["1"]
    );
}

#[test]
fn math_cos_of_zero_is_one() {
    assert_eq!(
        run_csharp(
            r#"double v = System.Math.Cos(0);
Console.WriteLine(v);"#
        ),
        &["1"]
    );
}

#[test]
fn math_tan_of_pi_over_four_approaches_one() {
    assert_eq!(
        run_csharp(
            r#"double v = System.Math.Tan(System.Math.PI / 4);
Console.WriteLine(System.Math.Round(v));"#
        ),
        &["1"]
    );
}

#[test]
fn math_truncate_removes_fractional_part_toward_zero() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.Truncate(9.9));"#),
        &["9"]
    );
}

#[test]
fn math_sign_returns_minus_one_for_negative() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.Sign(-42));"#),
        &["-1"]
    );
}

#[test]
fn math_min_returns_smaller_of_two_values() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.Min(3, 7));"#),
        &["3"]
    );
}

#[test]
fn math_pi_constant_has_correct_first_digits() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.PI > 3.14 && System.Math.PI < 3.15);"#),
        &["True"]
    );
}

#[test]
fn math_e_constant_has_correct_first_digits() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.E > 2.71 && System.Math.E < 2.72);"#),
        &["True"]
    );
}

#[test]
fn math_atan2_computes_angle_from_y_x_coordinates() {
    assert_eq!(
        run_csharp(
            r#"double angle = System.Math.Atan2(1, 1);
Console.WriteLine(System.Math.Round(angle, 4));"#
        ),
        &["0.7854"]
    );
}
