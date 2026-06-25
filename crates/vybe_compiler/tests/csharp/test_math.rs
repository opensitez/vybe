use super::helpers::{run_csharp, run_csharp_one};

#[test]
fn math_floor() {
    assert_eq!(run_csharp_one("Console.WriteLine(Math.Floor(3.7));"), "3");
}

#[test]
fn math_abs() {
    assert_eq!(run_csharp_one("Console.WriteLine(Math.Abs(-5));"), "5");
}

#[test]
fn math_sqrt() {
    assert_eq!(run_csharp_one("Console.WriteLine(Math.Sqrt(16));"), "4");
}

#[test]
fn math_multiple() {
    let out = run_csharp(
        r#"
        Console.WriteLine(Math.Floor(9.7));
        Console.WriteLine(Math.Abs(-42));
        Console.WriteLine(Math.Sqrt(144));
    "#,
    );
    assert_eq!(out, vec!["9", "42", "12"]);
}

#[test]
fn math_ceiling_rounds_positive_fraction_upward() {
    assert_eq!(run_csharp_one("Console.WriteLine(System.Math.Ceiling(2.1));"), "3");
}

#[test]
fn math_round_midpoint_to_even_for_half_values() {
    assert_eq!(run_csharp_one("Console.WriteLine(System.Math.Round(2.5));"), "2");
}

#[test]
fn math_max_selects_larger_of_two_doubles() {
    assert_eq!(run_csharp_one("Console.WriteLine(System.Math.Max(1.5, 2.5));"), "2.5");
}

#[test]
fn math_clamp_restricts_value_to_inclusive_bounds() {
    assert_eq!(run_csharp_one("Console.WriteLine(System.Math.Clamp(10, 0, 5));"), "5");
}

#[test]
fn math_div_rem_returns_quotient_and_remainder() {
    assert_eq!(
        run_csharp(
            r#"
int remainder;
var quotient = System.Math.DivRem(17, 5, out remainder);
Console.WriteLine(quotient);
Console.WriteLine(remainder);
"#
        ),
        &["3", "2"]
    );
}
