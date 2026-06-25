use crate::helpers::{run_python_one, run_print};

#[test]
fn arithmetic_add_integers() {
    assert_eq!(run_print("2 + 3"), "5");
}

#[test]
fn arithmetic_subtract() {
    assert_eq!(run_print("10 - 4"), "6");
}

#[test]
fn arithmetic_multiply() {
    assert_eq!(run_print("6 * 7"), "42");
}

#[test]
fn arithmetic_true_divide() {
    assert_eq!(run_print("7 / 2"), "3.5");
}

#[test]
fn arithmetic_floor_divide() {
    assert_eq!(run_print("7 // 2"), "3");
}

#[test]
fn arithmetic_modulo() {
    assert_eq!(run_print("10 % 3"), "1");
}

#[test]
fn arithmetic_power() {
    assert_eq!(run_print("2 ** 5"), "32");
}

#[test]
fn arithmetic_negation() {
    assert_eq!(run_print("-5"), "-5");
}

#[test]
fn arithmetic_unary_plus() {
    assert_eq!(run_print("+5"), "5");
}

#[test]
fn arithmetic_chained_add_sub() {
    assert_eq!(run_print("1 + 2 - 3"), "0");
}

#[test]
fn arithmetic_mul_before_add() {
    assert_eq!(run_print("2 + 3 * 4"), "14");
}

#[test]
fn arithmetic_parentheses_override() {
    assert_eq!(run_print("(2 + 3) * 4"), "20");
}

#[test]
fn arithmetic_float_add() {
    assert_eq!(run_print("0.1 + 0.2"), "0.30000000000000004");
}

#[test]
fn arithmetic_mixed_int_float() {
    assert_eq!(run_print("3 + 0.5"), "3.5");
}

#[test]
fn arithmetic_div_by_zero_raises() {
    assert_eq!(
        run_python_one("try:\n print(1/0)\nexcept ZeroDivisionError:\n print('z')\n"),
        "z"
    );
}

#[test]
fn arithmetic_floor_div_negative() {
    assert_eq!(run_print("-7 // 2"), "-4");
}

#[test]
fn arithmetic_mod_negative() {
    assert_eq!(run_print("-7 % 3"), "2");
}

#[test]
fn arithmetic_power_zero_exp() {
    assert_eq!(run_print("9 ** 0"), "1");
}

#[test]
fn arithmetic_power_one_exp() {
    assert_eq!(run_print("9 ** 1"), "9");
}

#[test]
fn arithmetic_large_int_add() {
    assert_eq!(run_print("10**10 + 1"), "10000000001");
}

#[test]
fn arithmetic_string_repeat() {
    assert_eq!(run_print("'ab' * 3"), "ababab");
}

#[test]
fn arithmetic_list_repeat() {
    assert_eq!(run_print("[0] * 3"), "[0, 0, 0]");
}

#[test]
fn arithmetic_tuple_repeat() {
    assert_eq!(run_print("(1,) * 3"), "(1, 1, 1)");
}

#[test]
fn arithmetic_in_place_add_list() {
    assert_eq!(
        run_python_one("a = [1]\na += [2]\nprint(a)\n"),
        "[1, 2]"
    );
}

#[test]
fn arithmetic_in_place_mul_list() {
    assert_eq!(
        run_python_one("a = [1]\na *= 3\nprint(a)\n"),
        "[1, 1, 1]"
    );
}

#[test]
fn arithmetic_complex_chain() {
    assert_eq!(run_print("1 + 2 * 3 - 4 // 2"), "5");
}

#[test]
fn arithmetic_divmod_builtin() {
    assert_eq!(run_print("divmod(7, 2)"), "(3, 1)");
}

#[test]
fn arithmetic_pow_builtin() {
    assert_eq!(run_print("pow(2, 3)"), "8");
}

#[test]
fn arithmetic_pow_mod_three_arg() {
    assert_eq!(run_print("pow(2, 3, 5)"), "3");
}

#[test]
fn arithmetic_abs_int() {
    assert_eq!(run_print("abs(-8)"), "8");
}

#[test]
fn arithmetic_round_half_up() {
    assert_eq!(run_print("round(2.5)"), "2");
}

#[test]
fn arithmetic_round_digits() {
    assert_eq!(run_print("round(1.234, 2)"), "1.23");
}

#[test]
fn arithmetic_int_from_float_trunc() {
    assert_eq!(run_print("int(3.9)"), "3");
}

#[test]
fn arithmetic_float_from_int() {
    assert_eq!(run_print("float(3)"), "3.0");
}

#[test]
fn arithmetic_bool_as_int_add() {
    assert_eq!(run_print("True + True"), "2");
}

#[test]
fn arithmetic_bool_multiply() {
    assert_eq!(run_print("False * 99"), "0");
}

#[test]
fn arithmetic_hex_literal_add() {
    assert_eq!(run_print("0x10 + 1"), "17");
}

#[test]
fn arithmetic_binary_literal_add() {
    assert_eq!(run_print("0b10 + 1"), "3");
}

#[test]
fn arithmetic_octal_literal_add() {
    assert_eq!(run_print("0o10 + 1"), "9");
}

#[test]
fn arithmetic_underscores_in_literal() {
    assert_eq!(run_print("1_000 + 2_000"), "3000");
}

#[test]
fn arithmetic_subtract_to_negative() {
    assert_eq!(run_print("3 - 10"), "-7");
}

#[test]
fn arithmetic_zero_div_float() {
    assert_eq!(run_print("0.0 / 5"), "0.0");
}

#[test]
fn arithmetic_inf_not_in_basic() {
    assert_eq!(run_print("1e308 * 2 > 1e308"), "True");
}

#[test]
fn arithmetic_complex_sum_parts() {
    assert_eq!(run_print("(1+2j) + (3+4j)"), "(4+6j)");
}

#[test]
fn arithmetic_expression_in_fstring() {
    assert_eq!(run_python_one("print(f'{3 * 4}')\n"), "12");
}

#[test]
fn arithmetic_sum_builtin() {
    assert_eq!(run_print("sum([1, 2, 3])"), "6");
}

#[test]
fn arithmetic_min_max() {
    assert_eq!(run_print("[min(1,2), max(1,2)]"), "[1, 2]");
}
