use crate::helpers::{run_print, run_python_one};

#[test]
fn pow_two_to_three() {
    assert_eq!(run_print("pow(2, 3)"), "8");
}

#[test]
fn pow_operator_two_stars() {
    assert_eq!(run_print("2 ** 3"), "8");
}

#[test]
fn pow_zero_exponent_is_one() {
    assert_eq!(run_print("5 ** 0"), "1");
}

#[test]
fn pow_one_exponent_is_base() {
    assert_eq!(run_print("99 ** 1"), "99");
}

#[test]
fn pow_zero_base_zero_is_one() {
    assert_eq!(run_print("0 ** 0"), "1");
}

#[test]
fn pow_negative_exponent_reciprocal() {
    assert_eq!(run_print("2 ** -1"), "0.5");
}

#[test]
fn pow_negative_base_even_exponent() {
    assert_eq!(run_print("(-2) ** 4"), "16");
}

#[test]
fn pow_negative_base_odd_exponent() {
    assert_eq!(run_print("(-2) ** 3"), "-8");
}

#[test]
fn pow_float_base() {
    assert_eq!(run_print("2.5 ** 2"), "6.25");
}

#[test]
fn pow_third_argument_modulo() {
    assert_eq!(run_print("pow(2, 10, 1000)"), "24");
}

#[test]
fn pow_modular_five_to_four_mod_thirteen() {
    assert_eq!(run_print("pow(5, 4, 13)"), "1");
}

#[test]
fn exponentiation_assign() {
    assert_eq!(run_python_one("n = 3\nn **= 4\nprint(n)\n"), "81");
}

#[test]
fn pow_in_expression_with_add() {
    assert_eq!(run_print("2 ** 3 + 1"), "9");
}

#[test]
fn pow_precedence_over_multiplication() {
    assert_eq!(run_print("2 * 3 ** 2"), "18");
}

#[test]
fn pow_right_associative_chain() {
    assert_eq!(run_print("2 ** 3 ** 2"), "512");
}

#[test]
fn pow_square_of_sum_in_parens() {
    assert_eq!(run_print("(2 + 3) ** 2"), "25");
}

#[test]
fn pow_list_comprehension_squares() {
    assert_eq!(run_print("[n ** 2 for n in range(5)]"), "[0, 1, 4, 9, 16]");
}

#[test]
fn pow_ten_to_positive_powers() {
    assert_eq!(run_print("10 ** 3"), "1000");
}

#[test]
fn pow_three_to_zero_through_four() {
    assert_eq!(run_print("[3 ** n for n in range(5)]"), "[1, 3, 9, 27, 81]");
}

#[test]
fn pow_modulo_large_composite() {
    assert_eq!(run_print("pow(7, 5, 11)"), "2");
}

#[test]
fn pow_fractional_exponent_sqrt() {
    assert_eq!(run_print("9 ** 0.5"), "3.0");
}

#[test]
fn pow_in_fstring() {
    assert_eq!(run_print("f'{3 ** 4}'"), "81");
}

#[test]
fn pow_nested_parens() {
    assert_eq!(run_print("(2 ** (3 ** 2))"), "512");
}

#[test]
fn pow_base_one_any_exponent() {
    assert_eq!(run_print("1 ** 100"), "1");
}

#[test]
fn pow_zero_base_positive_exponent() {
    assert_eq!(run_print("0 ** 5"), "0");
}

#[test]
fn pow_negative_exponent_creates_float() {
    assert_eq!(run_print("10 ** -2"), "0.01");
}

#[test]
fn pow_modulo_result_less_than_modulus() {
    assert_eq!(
        run_python_one("r = pow(123, 7, 50)\nprint(r < 50)\n"),
        "True"
    );
}

#[test]
fn pow_in_while_loop_counter() {
    assert_eq!(
        run_python_one("n = 0\nv = 1\nwhile n < 4:\n v = v * 2\n n += 1\nprint(v)\n"),
        "16"
    );
}

#[test]
fn pow_two_raised_to_ten() {
    assert_eq!(run_print("2 ** 10"), "1024");
}

#[test]
fn pow_five_cubed() {
    assert_eq!(run_print("5 ** 3"), "125");
}

#[test]
fn pow_seven_squared() {
    assert_eq!(run_print("7 ** 2"), "49");
}

#[test]
fn pow_modulo_identity_base_mod_one() {
    assert_eq!(run_print("pow(99, 5, 1)"), "0");
}

#[test]
fn pow_operator_with_unary_minus_on_result() {
    assert_eq!(run_print("-(2 ** 3)"), "-8");
}

#[test]
fn pow_sum_of_powers() {
    assert_eq!(run_print("2 ** 0 + 2 ** 1 + 2 ** 2"), "7");
}

#[test]
fn pow_in_dict_comprehension() {
    assert_eq!(
        run_print("{n: n ** 2 for n in range(4)}"),
        "{0: 0, 1: 1, 2: 4, 3: 9}"
    );
}

#[test]
fn pow_geometric_growth_check() {
    assert_eq!(
        run_python_one("a = 1\nfor _ in range(5):\n a *= 3\nprint(a)\n"),
        "243"
    );
}

#[test]
fn pow_fourth_power_of_three() {
    assert_eq!(run_print("3 ** 4"), "81");
}

#[test]
fn pow_sixth_power_of_two() {
    assert_eq!(run_print("2 ** 6"), "64");
}

#[test]
fn pow_eighth_power_of_two() {
    assert_eq!(run_print("2 ** 8"), "256");
}

#[test]
fn pow_modular_inverse_style() {
    assert_eq!(run_print("pow(3, 3, 7)"), "6");
}

#[test]
fn pow_expression_in_boolean_context() {
    assert_eq!(run_print("bool(2 ** 0)"), "False");
}

#[test]
fn pow_expression_truthy_when_positive() {
    assert_eq!(run_print("bool(2 ** 5)"), "True");
}

#[test]
fn pow_float_to_integer_power() {
    assert_eq!(run_print("1.5 ** 3"), "3.375");
}

#[test]
fn pow_negative_float_base_even() {
    assert_eq!(run_print("(-1.5) ** 2"), "2.25");
}

#[test]
fn pow_in_tuple_literal() {
    assert_eq!(run_print("(2 ** 3, 3 ** 2)"), "(8, 9)");
}

#[test]
fn pow_modulo_preserves_parity_check() {
    assert_eq!(run_print("pow(2, 10, 2)"), "0");
}

#[test]
fn pow_modulo_odd_base() {
    assert_eq!(run_print("pow(3, 4, 5)"), "1");
}

#[test]
fn pow_combined_with_floor_div() {
    assert_eq!(run_print("(2 ** 10) // 128"), "8");
}

#[test]
fn pow_combined_with_modulo() {
    assert_eq!(run_print("(2 ** 10) % 100"), "24");
}
