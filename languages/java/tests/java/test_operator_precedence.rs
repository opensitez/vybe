use crate::helpers::run_main;

#[test]
fn multiplication_binds_tighter_than_addition() {
    let out = run_main("System.out.println(2 + 3 * 4);");
    assert_eq!(out, vec!["14"]);
}

#[test]
fn division_binds_tighter_than_subtraction() {
    let out = run_main("System.out.println(10 - 8 / 2);");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn modulo_same_precedence_as_multiply_left_to_right() {
    let out = run_main("System.out.println(20 - 10 % 3 * 2);");
    assert_eq!(out, vec!["16"]);
}

#[test]
fn chained_multiplication_and_division_left_to_right() {
    let out = run_main("System.out.println(24 / 4 * 2);");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn chained_addition_and_subtraction_left_to_right() {
    let out = run_main("System.out.println(10 - 3 + 2);");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn parentheses_force_addition_before_multiplication() {
    let out = run_main("System.out.println((2 + 3) * 4);");
    assert_eq!(out, vec!["20"]);
}

#[test]
fn nested_parentheses_group_inner_sum_first() {
    let out = run_main("System.out.println(((1 + 2) * 3) + 4);");
    assert_eq!(out, vec!["13"]);
}

#[test]
fn parentheses_override_modulo_and_division_order() {
    let out = run_main("System.out.println((10 % 3) * (8 / 2));");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn unary_minus_binds_tighter_than_multiplication() {
    let out = run_main("System.out.println(-3 * 4);");
    assert_eq!(out, vec!["-12"]);
}

#[test]
fn unary_minus_on_parenthesized_expression() {
    let out = run_main("System.out.println(-(3 + 2));");
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn unary_minus_with_addition_left_to_right() {
    let out = run_main("System.out.println(5 - -2);");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn double_unary_minus_negates_twice() {
    let out = run_main("System.out.println(--5);");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn unary_plus_is_noop_on_literal() {
    let out = run_main("System.out.println(+7 + 3);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn prefix_increment_before_multiplication() {
    let out = run_main("int x = 4; System.out.println(++x * 2);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn postfix_increment_after_multiplication_uses_old_value() {
    let out = run_main("int x = 4; System.out.println(x++ * 2);");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn prefix_decrement_in_subtraction_expression() {
    let out = run_main("int x = 5; System.out.println(x - --x);");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn postfix_increment_in_addition_chain() {
    let out = run_main("int x = 1; System.out.println(x++ + x + x++);");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn int_plus_double_promotes_to_double() {
    let out = run_main("System.out.println(1 + 2.0 * 3);");
    assert_eq!(out, vec!["7.0"]);
}

#[test]
fn int_division_before_double_addition() {
    let out = run_main("System.out.println(5 / 2 + 0.5);");
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn double_literal_forces_floating_division() {
    let out = run_main("System.out.println(5 / 2.0);");
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn mixed_int_double_subtraction_promotes() {
    let out = run_main("System.out.println(10.0 - 3);");
    assert_eq!(out, vec!["7.0"]);
}

#[test]
fn int_multiplied_by_double_becomes_double() {
    let out = run_main("System.out.println(3 * 2.5);");
    assert_eq!(out, vec!["7.5"]);
}

#[test]
fn parentheses_with_mixed_types_inside() {
    let out = run_main("System.out.println((2 + 3) * 1.0);");
    assert_eq!(out, vec!["5.0"]);
}

#[test]
fn exponent_style_via_multiplication_chain() {
    let out = run_main("System.out.println(2 * 2 * 2 * 2);");
    assert_eq!(out, vec!["16"]);
}

#[test]
fn negation_of_product_requires_parentheses_for_positive() {
    let out = run_main("System.out.println(-(2 * 3));");
    assert_eq!(out, vec!["-6"]);
}

#[test]
fn addition_of_negative_literals() {
    let out = run_main("System.out.println(-3 + -4);");
    assert_eq!(out, vec!["-7"]);
}

#[test]
fn subtraction_with_parenthesized_sum() {
    let out = run_main("System.out.println(20 - (5 + 3));");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn modulo_before_addition_without_parens() {
    let out = run_main("System.out.println(7 % 4 + 10);");
    assert_eq!(out, vec!["13"]);
}

#[test]
fn division_and_modulo_left_associative() {
    let out = run_main("System.out.println(17 % 5 / 2);");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn complex_parenthesis_with_three_operators() {
    let out = run_main("System.out.println((6 + 2) * (5 - 3) - 4);");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn unary_minus_on_variable_times_positive() {
    let out = run_main("int x = 5; System.out.println(-x * 2);");
    assert_eq!(out, vec!["-10"]);
}

#[test]
fn variable_plus_prefix_increment_times_two() {
    let out = run_main("int x = 3; System.out.println(x + ++x);");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn double_modulo_with_double_addition() {
    let out = run_main("System.out.println(5.5 % 2.0 + 1.0);");
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn int_expression_in_nested_parens() {
    let out = run_main("System.out.println(1 + (2 + (3 + 4)));");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn multiply_add_mix_with_double_operand() {
    let out = run_main("System.out.println(2.0 + 3 * 4);");
    assert_eq!(out, vec!["14.0"]);
}

#[test]
fn subtract_divide_mix_respects_precedence() {
    let out = run_main("System.out.println(9.0 - 6 / 2);");
    assert_eq!(out, vec!["6.0"]);
}

#[test]
fn long_arithmetic_stays_integral() {
    let out = run_main("System.out.println(10L + 5L * 2L);");
    assert_eq!(out, vec!["20"]);
}

#[test]
fn cast_after_arithmetic_before_division() {
    let out = run_main("System.out.println((int) (7 + 3) / 2);");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn negated_sum_of_products() {
    let out = run_main("System.out.println(-(2 * 3 + 4));");
    assert_eq!(out, vec!["-10"]);
}

#[test]
fn deeply_parenthesized_mixed_int_double() {
    let out = run_main("System.out.println(((1 + 2.0) * 2) + 1);");
    assert_eq!(out, vec!["7.0"]);
}
