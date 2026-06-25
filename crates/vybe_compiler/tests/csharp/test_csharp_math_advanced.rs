//! Less-common `System.Math` and `System.MathF` members.
use super::helpers::run_csharp;

#[test]
fn math_abs_for_negative_double() {
    assert_eq!(run_csharp(r#"Console.WriteLine(System.Math.Abs(-3.7));"#), &["3.7"]);
}

#[test]
fn math_floor_rounds_toward_negative_infinity() {
    assert_eq!(run_csharp(r#"Console.WriteLine(System.Math.Floor(2.9));"#), &["2"]);
}

#[test]
fn math_ceiling_rounds_toward_positive_infinity() {
    assert_eq!(run_csharp(r#"Console.WriteLine(System.Math.Ceiling(2.1));"#), &["3"]);
}

#[test]
fn math_round_midpoint_away_from_zero() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.Round(2.5,System.MidpointRounding.AwayFromZero));"#),
        &["3"]
    );
}

#[test]
fn math_round_midpoint_to_even_banker_rounding() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.Round(2.5,System.MidpointRounding.ToEven));"#),
        &["2"]
    );
}

#[test]
fn math_log2_of_power_of_two() {
    assert_eq!(run_csharp(r#"Console.WriteLine((int)System.Math.Log2(8));"#), &["3"]);
}

#[test]
fn math_bit_decrement_returns_next_lower_double() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.BitDecrement(1.0)<1.0);"#),
        &["True"]
    );
}

#[test]
fn math_clamp_returns_low_when_value_below_range() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.Clamp(-5,0,10));"#),
        &["0"]
    );
}

#[test]
fn math_clamp_returns_value_when_in_range() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Math.Clamp(5,0,10));"#),
        &["5"]
    );
}
