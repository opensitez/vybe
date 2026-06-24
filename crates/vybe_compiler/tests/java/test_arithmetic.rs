use crate::helpers::run_main;

#[test]
fn int_addition_sums_operands() {
    let out = run_main("System.out.println(3 + 4);");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn int_subtraction_finds_difference() {
    let out = run_main("System.out.println(10 - 3);");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn int_multiplication_scales_value() {
    let out = run_main("System.out.println(6 * 7);");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn int_division_truncates_toward_zero() {
    let out = run_main("System.out.println(7 / 2);");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn int_modulo_returns_remainder() {
    let out = run_main("System.out.println(17 % 5);");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn int_modulo_keeps_dividend_sign_for_negative() {
    let out = run_main("System.out.println(-17 % 5);");
    assert_eq!(out, vec!["-2"]);
}

#[test]
fn int_overflow_wraps_past_max_value() {
    let out = run_main("System.out.println(Integer.MAX_VALUE + 1);");
    assert_eq!(out, vec!["-2147483648"]);
}

#[test]
fn int_underflow_wraps_below_min_value() {
    let out = run_main("System.out.println(Integer.MIN_VALUE - 1);");
    assert_eq!(out, vec!["2147483647"]);
}

#[test]
fn long_addition_preserves_wide_sum() {
    let out = run_main("System.out.println(1_000_000L + 2_000_000L);");
    assert_eq!(out, vec!["3000000"]);
}

#[test]
fn long_subtraction_preserves_wide_difference() {
    let out = run_main("System.out.println(9_000_000L - 1_000_000L);");
    assert_eq!(out, vec!["8000000"]);
}

#[test]
fn long_multiplication_scales_wide_value() {
    let out = run_main("System.out.println(1000L * 1000L);");
    assert_eq!(out, vec!["1000000"]);
}

#[test]
fn long_division_truncates_toward_zero() {
    let out = run_main("System.out.println(9L / 2L);");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn long_modulo_returns_remainder() {
    let out = run_main("System.out.println(19L % 4L);");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn double_addition_sums_fractions() {
    let out = run_main("System.out.println(1.5 + 2.5);");
    assert_eq!(out, vec!["4.0"]);
}

#[test]
fn double_subtraction_finds_difference() {
    let out = run_main("System.out.println(5.5 - 2.0);");
    assert_eq!(out, vec!["3.5"]);
}

#[test]
fn double_multiplication_scales_value() {
    let out = run_main("System.out.println(2.5 * 4.0);");
    assert_eq!(out, vec!["10.0"]);
}

#[test]
fn double_division_divides_fractions() {
    let out = run_main("System.out.println(7.5 / 2.5);");
    assert_eq!(out, vec!["3.0"]);
}

#[test]
fn double_modulo_returns_fractional_remainder() {
    let out = run_main("System.out.println(5.5 % 2.0);");
    assert_eq!(out, vec!["1.5"]);
}

#[test]
fn compound_addition_assignment_accumulates() {
    let out = run_main("int x = 10; x += 5; System.out.println(x);");
    assert_eq!(out, vec!["15"]);
}

#[test]
fn compound_subtraction_assignment_reduces() {
    let out = run_main("int x = 10; x -= 3; System.out.println(x);");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn compound_multiplication_assignment_scales() {
    let out = run_main("int x = 6; x *= 7; System.out.println(x);");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn compound_division_assignment_quotients() {
    let out = run_main("int x = 20; x /= 4; System.out.println(x);");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn compound_modulo_assignment_remainders() {
    let out = run_main("int x = 17; x %= 5; System.out.println(x);");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn prefix_increment_updates_before_use() {
    let out = run_main("int x = 5; System.out.println(++x); System.out.println(x);");
    assert_eq!(out, vec!["6", "6"]);
}

#[test]
fn postfix_increment_uses_old_value_then_updates() {
    let out = run_main("int x = 5; System.out.println(x++); System.out.println(x);");
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn prefix_decrement_updates_before_use() {
    let out = run_main("int x = 5; System.out.println(--x); System.out.println(x);");
    assert_eq!(out, vec!["4", "4"]);
}

#[test]
fn postfix_decrement_uses_old_value_then_updates() {
    let out = run_main("int x = 5; System.out.println(x--); System.out.println(x);");
    assert_eq!(out, vec!["5", "4"]);
}

#[test]
fn prefix_increment_in_expression_uses_new_value() {
    let out = run_main("int x = 4; System.out.println(++x + 1);");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn postfix_increment_in_expression_uses_old_value() {
    let out = run_main("int x = 4; System.out.println(x++ + 1);");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn long_prefix_increment_advances_wide_counter() {
    let out = run_main("long x = 9L; System.out.println(++x);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn long_postfix_decrement_preserves_old_wide_value() {
    let out = run_main("long x = 9L; System.out.println(x--); System.out.println(x);");
    assert_eq!(out, vec!["9", "8"]);
}

#[test]
fn double_prefix_increment_advances_fractional_value() {
    let out = run_main("double x = 1.5; System.out.println(++x);");
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn double_postfix_decrement_preserves_old_fractional_value() {
    let out = run_main("double x = 2.5; System.out.println(x--); System.out.println(x);");
    assert_eq!(out, vec!["2.5", "1.5"]);
}

#[test]
fn unary_negation_flips_sign_on_int() {
    let out = run_main("int x = 8; System.out.println(-x);");
    assert_eq!(out, vec!["-8"]);
}

#[test]
fn unary_negation_flips_sign_on_double() {
    let out = run_main("double x = 3.5; System.out.println(-x);");
    assert_eq!(out, vec!["-3.5"]);
}

#[test]
fn int_expression_respects_multiplication_before_addition() {
    let out = run_main("System.out.println(2 + 3 * 4);");
    assert_eq!(out, vec!["14"]);
}

#[test]
fn parentheses_override_default_precedence() {
    let out = run_main("System.out.println((2 + 3) * 4);");
    assert_eq!(out, vec!["20"]);
}

#[test]
fn chained_addition_accumulates_left_to_right() {
    let out = run_main("System.out.println(1 + 2 + 3 + 4);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn increment_on_array_element_updates_slot() {
    let out = run_main("int[] arr = {1, 2}; arr[0]++; System.out.println(arr[0]);");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn decrement_on_array_element_updates_slot() {
    let out = run_main("int[] arr = {1, 2}; --arr[1]; System.out.println(arr[1]);");
    assert_eq!(out, vec!["1"]);
}
