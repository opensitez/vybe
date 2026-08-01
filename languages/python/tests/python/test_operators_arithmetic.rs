use crate::helpers::{run_print, run_python_one};

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
    assert_eq!(run_python_one("a = [1]\na += [2]\nprint(a)\n"), "[1, 2]");
}

#[test]
fn arithmetic_in_place_mul_list() {
    assert_eq!(run_python_one("a = [1]\na *= 3\nprint(a)\n"), "[1, 1, 1]");
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

// ── `%` is FLOORED — the result takes the DIVISOR's sign ───────────────────
//
// `[builtin_slots.int] mod = "common:math.floor_mod"` (builtinslotplan.md
// §3i). The platform default truncates, as C and JS do, so `-7 % 3` would be
// -1 there and is 2 here. `-7 % 3` was already covered; these pin the negative
// DIVISOR (where truncation and flooring differ in the other direction) and the
// float operands, which take the same path.

#[test]
fn modulo_negative_divisor_takes_divisor_sign() {
    assert_eq!(run_print("7 % -3"), "-2");
}

#[test]
fn modulo_both_negative() {
    assert_eq!(run_print("-7 % -3"), "-1");
}

#[test]
fn modulo_float_operands_floors() {
    assert_eq!(run_print("7.5 % 2"), "1.5");
}

#[test]
fn modulo_negative_float_takes_divisor_sign() {
    assert_eq!(run_print("-7.5 % 2"), "0.5");
}

// ── `len()` counts CODE POINTS, not UTF-16 code units ──────────────────────
//
// `unifiedstringplan.md` Axis 1 names Python's index unit `scalar`, and Python
// is the only language on that axis. The shared length helper counts UTF-16
// units, so every character outside the BMP was off by one: `len("😀")` was 2
// and `len("a😀b")` was 4.
//
// `Array.from` walks a string with the STRING ITERATOR, which yields code
// points, so its length is the scalar count — the `[...s].length` idiom over
// primitives that already existed. No host function was added.
//
// Values measured against real `python3`.

#[test]
fn len_of_non_bmp_char_is_one_code_point() {
    assert_eq!(run_print(r#"len("\U0001F600")"#), "1");
}

#[test]
fn len_counts_code_points_around_a_non_bmp_char() {
    assert_eq!(run_print(r#"len("a\U0001F600b")"#), "3");
}

#[test]
fn len_of_accented_string_is_unchanged() {
    assert_eq!(run_print(r#"len("café")"#), "4");
}

#[test]
fn len_of_empty_string_is_zero() {
    assert_eq!(run_print(r#"len("")"#), "0");
}

// `len` is polymorphic — the string leg must not disturb the others.
#[test]
fn len_of_list_is_element_count() {
    assert_eq!(run_print("len([1, 2, 3])"), "3");
}

#[test]
fn len_of_dict_is_key_count() {
    assert_eq!(run_print("len({'a': 1})"), "1");
}

#[test]
fn len_of_set_is_element_count() {
    assert_eq!(run_print("len({1, 2})"), "2");
}

#[test]
fn len_of_bytes_is_byte_count() {
    assert_eq!(run_print(r#"len(b"abc")"#), "3");
}

#[test]
fn len_uses_user_dunder_len() {
    assert_eq!(
        run_print("class C:\n    def __len__(s): return 7\nlen(C())"),
        "7"
    );
}
