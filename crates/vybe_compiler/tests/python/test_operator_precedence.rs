use crate::helpers::{run_python_one, run_print};

#[test]
fn precedence_mul_over_add() {
    assert_eq!(run_print("2 + 3 * 4"), "14");
}

#[test]
fn precedence_paren_over_mul() {
    assert_eq!(run_print("(2 + 3) * 4"), "20");
}

#[test]
fn precedence_power_over_mul() {
    assert_eq!(run_print("2 * 3 ** 2"), "18");
}

#[test]
fn precedence_power_right_associative() {
    assert_eq!(run_print("2 ** 3 ** 2"), "512");
}

#[test]
fn precedence_unary_minus_over_power() {
    assert_eq!(run_print("-2 ** 2"), "-4");
}

#[test]
fn precedence_unary_minus_paren() {
    assert_eq!(run_print("(-2) ** 2"), "4");
}

#[test]
fn precedence_floor_div_same_as_mul() {
    assert_eq!(run_print("10 - 8 // 2"), "6");
}

#[test]
fn precedence_mod_same_as_mul() {
    assert_eq!(run_print("10 - 9 % 4"), "7");
}

#[test]
fn precedence_comparison_over_and() {
    assert_eq!(run_print("1 < 2 and 3 < 4"), "True");
}

#[test]
fn precedence_and_over_or() {
    assert_eq!(run_print("False or True and False"), "False");
}

#[test]
fn precedence_not_over_and() {
    assert_eq!(run_print("not False and True"), "True");
}

#[test]
fn precedence_bitwise_and_over_xor() {
    assert_eq!(run_print("5 ^ 3 & 6"), "3");
}

#[test]
fn precedence_bitwise_xor_over_or() {
    assert_eq!(run_print("5 | 3 ^ 6"), "7");
}

#[test]
fn precedence_bitwise_over_comparison() {
    assert_eq!(run_print("5 & 3 > 1"), "True");
}

#[test]
fn precedence_shift_over_add() {
    assert_eq!(run_print("1 + 2 << 2"), "12");
}

#[test]
fn precedence_chained_comparison() {
    assert_eq!(run_print("1 < 2 < 3"), "True");
}

#[test]
fn precedence_chained_comparison_false() {
    assert_eq!(run_print("1 < 3 < 2"), "False");
}

#[test]
fn precedence_chained_with_equal() {
    assert_eq!(run_print("1 < 2 == True"), "True");
}

#[test]
fn precedence_member_over_comparison() {
    assert_eq!(run_print("1 in [1, 2] == True"), "True");
}

#[test]
fn precedence_attr_over_call_not_applicable() {
    assert_eq!(run_print("len([1, 2])"), "2");
}

#[test]
fn precedence_subscript_over_power() {
    assert_eq!(run_print("['a', 'b'][1] ** 2"), "1");
}

#[test]
fn precedence_call_over_power() {
    assert_eq!(
        run_python_one("def f():\n return 2\nprint(f() ** 3)\n"),
        "8"
    );
}

#[test]
fn precedence_conditional_over_or() {
    assert_eq!(run_print("1 or 2 if False else 3"), "1");
}

#[test]
fn precedence_conditional_low() {
    assert_eq!(run_print("3 if 1 < 2 else 4"), "3");
}

#[test]
fn precedence_lambda_body_binds_tight() {
    assert_eq!(
        run_python_one("f = lambda: 1 + 2 * 3\nprint(f())\n"),
        "7"
    );
}

#[test]
fn precedence_list_comp_over_conditional() {
    assert_eq!(
        run_print("[x for x in range(3) if x]"),
        "[1, 2]"
    );
}

#[test]
fn precedence_star_unpack_low_in_call() {
    assert_eq!(
        run_python_one("def f(a, b):\n return a + b\nprint(f(*[1, 2]))\n"),
        "3"
    );
}

#[test]
fn precedence_walrus_over_conditional() {
    assert_eq!(
        run_python_one("print(1 if (n := 2) else 0)\n"),
        "1"
    );
}

#[test]
fn precedence_bool_ops_short_circuit() {
    assert_eq!(
        run_python_one("def boom():\n raise ValueError\nprint(0 and boom())\n"),
        "0"
    );
}

#[test]
fn precedence_or_short_circuit() {
    assert_eq!(
        run_python_one("def boom():\n raise ValueError\nprint(1 or boom())\n"),
        "1"
    );
}

#[test]
fn precedence_mixed_arithmetic_left_to_right() {
    assert_eq!(run_print("10 - 3 - 2"), "5");
}

#[test]
fn precedence_mixed_div_mul_left_to_right() {
    assert_eq!(run_print("8 / 2 * 2"), "8.0");
}

#[test]
fn precedence_float_true_div() {
    assert_eq!(run_print("1 / 2 * 2"), "1.0");
}

#[test]
fn precedence_nested_parens() {
    assert_eq!(run_print("((1 + 2) * (3 + 4))"), "21");
}

#[test]
fn precedence_compare_with_arithmetic() {
    assert_eq!(run_print("2 + 2 == 4"), "True");
}

#[test]
fn precedence_compare_with_arithmetic_false() {
    assert_eq!(run_print("2 + 2 == 5"), "False");
}

#[test]
fn precedence_in_not_in() {
    assert_eq!(run_print("3 not in [1, 2, 3]"), "False");
}

#[test]
fn precedence_is_not_over_and() {
    assert_eq!(run_print("[] is not None and True"), "True");
}

#[test]
fn precedence_multiple_and() {
    assert_eq!(run_print("True and True and False"), "False");
}

#[test]
fn precedence_multiple_or() {
    assert_eq!(run_print("False or False or 7"), "7");
}

#[test]
fn precedence_not_not() {
    assert_eq!(run_print("not not 0"), "True");
}

#[test]
fn precedence_complex_bitwise_mixed() {
    assert_eq!(run_print("~0 & 1"), "1");
}

#[test]
fn precedence_shift_before_and() {
    assert_eq!(run_print("1 << 2 & 7"), "4");
}

#[test]
fn precedence_add_before_shift() {
    assert_eq!(run_print("1 + 2 << 1"), "6");
}

#[test]
fn precedence_subscription_on_call_result() {
    assert_eq!(
        run_python_one("def f():\n return [1, 2]\nprint(f()[0])\n"),
        "1"
    );
}

#[test]
fn precedence_method_call_chain() {
    assert_eq!(run_print("'  hi  '.strip().upper()"), "HI");
}
