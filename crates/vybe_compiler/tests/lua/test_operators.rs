use super::helpers::{parse_ok, run_lua_one};

#[test]
fn addition() {
    let out = run_lua_one("print(1 + 2)\n");
    assert_eq!(out, "3");
}

#[test]
fn subtraction() {
    let out = run_lua_one("print(10 - 4)\n");
    assert_eq!(out, "6");
}

#[test]
fn multiplication() {
    let out = run_lua_one("print(3 * 7)\n");
    assert_eq!(out, "21");
}

#[test]
fn division() {
    let out = run_lua_one("print(8 / 2)\n");
    assert_eq!(out, "4");
}

#[test]
fn modulo() {
    let out = run_lua_one("print(10 % 3)\n");
    assert_eq!(out, "1");
}

#[test]
fn exponentiation() {
    let out = run_lua_one("print(2 ^ 3)\n");
    assert_eq!(out, "8");
}

#[test]
fn string_concat() {
    let out = run_lua_one("print(\"foo\" .. \"bar\")\n");
    assert_eq!(out, "foobar");
}

#[test]
fn equality() {
    let out = run_lua_one("print(1 == 1)\n");
    assert_eq!(out, "true");
}

#[test]
fn inequality() {
    let out = run_lua_one("print(1 ~= 2)\n");
    assert_eq!(out, "true");
}

#[test]
fn logical_and_short_circuit_shape() {
    parse_ok("local x = false and print(\"skip\")\n");
}

#[test]
fn logical_or() {
    let out = run_lua_one("print(false or 99)\n");
    assert_eq!(out, "99");
}

#[test]
fn unary_not() {
    let out = run_lua_one("print(not false)\n");
    assert_eq!(out, "true");
}

#[test]
fn unary_negation() {
    let out = run_lua_one("print(-5)\n");
    assert_eq!(out, "-5");
}

// ── Spec gaps: arithmetic / relational / logical (Lua 5.x manual §3.4) ─────

lua_print! {
    floor_division_truncates => { "print(7 // 3)\n", "2" },
    floor_division_negative_dividend => { "print(-7 // 3)\n", "-3" },
    power_binds_right_to_left => { "print(2 ^ 3 ^ 2)\n", "512" },
    multiplication_before_addition => { "print(2 + 3 * 4)\n", "14" },
    division_yields_float => { "print(5 / 2)\n", "2.5" },
    string_less_than_lexicographic => { "print(\"a\" < \"b\")\n", "true" },
    nil_equals_nil => { "print(nil == nil)\n", "true" },
    nil_not_equal_to_false => { "print(nil ~= false)\n", "true" },
    not_nil_is_true => { "print(not nil)\n", "true" },
    and_returns_last_when_both_truthy => { "print(10 and 20)\n", "20" },
    or_returns_first_truthy_operand => { "print(5 or 99)\n", "5" },
    and_or_idiom_selects_default => { "print(nil and 99 or 7)\n", "7" },
    zero_is_truthy_in_expression => { "print(0 and 99 or 7)\n", "99" },
    string_number_equality_does_not_coerce => { "print(\"10\" == 10)\n", "false" },
    string_number_less_than_does_not_coerce => { "print(\"2\" < 12)\n", "false" },
    concat_has_lower_precedence_than_arithmetic => { "print(\"x\" .. 1 + 1)\n", "x2" },
    arithmetic_has_higher_precedence_than_concat => { "print(1 + 2 .. \"!\")\n", "3!" },
    unary_minus_binds_tighter_than_power_without_parens => { "print(-2 ^ 2)\n", "4" },
    parentheses_force_unary_minus_before_power => { "print((-2) ^ 2)\n", "4" },
    false_and_true_short_circuits_to_false => { "print(false and true)\n", "false" },
    true_or_false_short_circuits_to_true => { "print(true or false)\n", "true" },
    floor_division_with_negative_divisor => { "print(-10 // -3)\n", "3" },
    modulo_result_sign_follows_divisor => { "print(-10 % 3)\n", "2" },
    table_and_string_comparison_returns_false => { "print({} == \"{}\")\n", "false" },
    different_tables_are_not_equal => { "print({} == {})\n", "false" },
    function_identity_compares_by_reference => {
        "local f = function() end\nprint(f == f)\n",
        "true"
    },
    and_returns_first_falsy_value => { "print(nil and 99)\n", "nil" },
    or_returns_first_truthy_value => { "print(0 or \"x\")\n", "0" },
    length_operator_on_string_in_expression => { "print(#\"abc\" == 3)\n", "true" },
    number_compared_to_numeric_string_coerces => { "print(1 < \"2\")\n", "true" },
    table_compared_to_number_is_false => { "print({} < 1)\n", "false" },
    boolean_compared_to_number_is_false => { "print(true < 2)\n", "false" },
    equality_false_between_different_types => { "print({} == 1)\n", "false" },
    less_than_on_numbers => { "print(2 < 5)\n", "true" },
    greater_than_on_numbers => { "print(9 > 4)\n", "true" },
    less_or_equal_when_equal => { "print(4 <= 4)\n", "true" },
    greater_or_equal_when_equal => { "print(7 >= 7)\n", "true" },
    inequality_on_strings => { "print(\"a\" ~= \"b\")\n", "true" },
    equality_on_strings => { "print(\"x\" == \"x\")\n", "true" },
    logical_not_on_true => { "print(not true)\n", "false" },
    logical_and_with_nil => { "print(nil and 1)\n", "nil" },
    logical_or_with_nil => { "print(nil or \"d\")\n", "d" },
    string_concat_three_parts => { "print(\"a\"..\"b\"..\"c\")\n", "abc" },
    length_of_literal_string => { "print(#\"abcd\")\n", "4" },
    negation_of_positive => { "print(-1)\n", "-1" },
    negation_of_negative => { "print(-(-3))\n", "3" },
    modulo_on_positive => { "print(17 % 5)\n", "2" },
    exponent_two_to_three => { "print(2 ^ 3)\n", "8" },
    division_produces_float => { "print(7 / 2)\n", "3.5" },
    subtract_to_zero => { "print(5 - 5)\n", "0" },
    multiply_by_one_is_identity => { "print(8 * 1)\n", "8" },
    add_negative_numbers => { "print(-2 + -3)\n", "-5" },
    zero_and_empty_string_are_truthy_in_and => {
        "print((0 and \"a\") == 0)\n",
        "0"
    },
    nil_or_returns_second_operand => {
        "print(nil or \"fallback\")\n",
        "fallback"
    },
    false_and_short_circuits_before_rhs_call => {
        "local n = 0\nlocal function bump() n = n + 1 return true end\nprint(false and bump())\n",
        "false"
    },
    true_or_short_circuits_before_rhs_call => {
        "local n = 0\nlocal function bump() n = n + 1 return true end\nprint(true or bump())\n",
        "true"
    },
    not_nil_is_true_value => { "print(not nil)\n", "true" },
    not_false_is_true_value => { "print(not false)\n", "true" },
    concatenation_left_associative_chain => { "print(\"a\"..\"b\"..\"c\"..\"d\")\n", "abcd" },
    comparison_chains_not_supported_use_two_ops => {
        "print(1 < 2 and 2 < 3)\n",
        "true"
    },
    equality_between_same_string_literals => { "print(\"lua\" == \"lua\")\n", "true" },
    inequality_between_number_and_string => { "print(1 ~= \"1\")\n", "true" },
}
