use crate::helpers::run_main;

#[test]
fn math_abs_negative_int() {
    let out = run_main("System.out.println(Math.abs(-15));");
    assert_eq!(out, vec!["15"]);
}

#[test]
fn math_abs_positive_int() {
    let out = run_main("System.out.println(Math.abs(9));");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn math_abs_zero() {
    let out = run_main("System.out.println(Math.abs(0));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn math_abs_negative_double() {
    let out = run_main("System.out.println(Math.abs(-2.5));");
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn math_min_picks_smaller_int() {
    let out = run_main("System.out.println(Math.min(3, 9));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn math_min_with_negative_operand() {
    let out = run_main("System.out.println(Math.min(-4, 2));");
    assert_eq!(out, vec!["-4"]);
}

#[test]
fn math_max_picks_larger_int() {
    let out = run_main("System.out.println(Math.max(3, 9));");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn math_max_equal_arguments_returns_either() {
    let out = run_main("System.out.println(Math.max(5, 5));");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn math_pow_squares_integer_base() {
    let out = run_main("System.out.println((int) Math.pow(3, 2));");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn math_pow_zero_exponent_yields_one() {
    let out = run_main("System.out.println((int) Math.pow(7, 0));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn math_sqrt_perfect_square() {
    let out = run_main("System.out.println((int) Math.sqrt(81));");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn math_sqrt_non_perfect_square_truncated() {
    let out = run_main("System.out.println((int) Math.sqrt(10));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn math_round_half_up_positive() {
    let out = run_main("System.out.println(Math.round(2.6));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn math_round_half_down_positive() {
    let out = run_main("System.out.println(Math.round(2.4));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn math_floor_positive_fraction() {
    let out = run_main("System.out.println(Math.floor(3.9));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn math_floor_negative_fraction() {
    let out = run_main("System.out.println(Math.floor(-1.2));");
    assert_eq!(out, vec!["-2"]);
}

#[test]
fn math_ceil_positive_fraction() {
    let out = run_main("System.out.println(Math.ceil(3.1));");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn math_ceil_negative_fraction() {
    let out = run_main("System.out.println(Math.ceil(-1.8));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn math_random_value_between_zero_and_one() {
    let out = run_main(
        "double r = Math.random(); System.out.println(r >= 0.0); System.out.println(r < 1.0);",
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn math_random_bounded_via_floor_multiply_pattern() {
    let out = run_main(
        "int bound = 10; int roll = (int) Math.floor(Math.random() * bound); System.out.println(roll >= 0); System.out.println(roll < bound);",
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn math_random_second_draw_produces_finite_value() {
    let out = run_main(
        "double a = Math.random(); double b = Math.random(); System.out.println(a >= 0.0); System.out.println(b >= 0.0);",
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn math_pi_constant() {
    let out = run_main("System.out.println(Math.PI > 3.14);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn math_e_constant() {
    let out = run_main("System.out.println(Math.E > 2.71);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn math_hypot_three_four_five_triangle() {
    let out = run_main("System.out.println((int) Math.hypot(3, 4));");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn math_signum_positive_value() {
    let out = run_main("System.out.println(Math.signum(12.0));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn math_signum_negative_value() {
    let out = run_main("System.out.println(Math.signum(-3.0));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn math_signum_zero() {
    let out = run_main("System.out.println(Math.signum(0.0));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn math_abs_on_negative_long_literal() {
    let out = run_main("System.out.println(Math.abs(-42L));");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn math_max_nested_call() {
    let out = run_main("System.out.println(Math.max(Math.max(1, 4), 3));");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn math_min_nested_call() {
    let out = run_main("System.out.println(Math.min(Math.min(8, 5), 6));");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn math_pow_fractional_exponent() {
    let out = run_main("System.out.println(Math.pow(4.0, 0.5));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn math_floor_div_positive_operands() {
    let out = run_main("System.out.println(Math.floorDiv(17, 5));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn math_floor_mod_positive_operands() {
    let out = run_main("System.out.println(Math.floorMod(17, 5));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn math_round_negative_fraction() {
    let out = run_main("System.out.println(Math.round(-2.6));");
    assert_eq!(out, vec!["-3"]);
}

#[test]
fn math_random_times_hundred_yields_bounded_range() {
    let out = run_main(
        "int n = (int) Math.floor(Math.random() * 100.0); System.out.println(n >= 0); System.out.println(n < 100);",
    );
    assert_eq!(out, vec!["true", "true"]);
}
