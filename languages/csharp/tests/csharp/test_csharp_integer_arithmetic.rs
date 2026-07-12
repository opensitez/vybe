use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(
    addition_of_two_positive_literals,
    r#"Console.WriteLine(12 + 8);"#,
    ["20"]
);

csharp_case!(
    addition_with_zero_as_first_operand,
    r#"Console.WriteLine(0 + 17);"#,
    ["17"]
);

csharp_case!(
    addition_with_zero_as_second_operand,
    r#"Console.WriteLine(17 + 0);"#,
    ["17"]
);

csharp_case!(
    addition_of_negative_and_positive_yields_difference,
    r#"Console.WriteLine(-5 + 12);"#,
    ["7"]
);

csharp_case!(
    addition_of_two_negatives_yields_more_negative_sum,
    r#"Console.WriteLine(-4 + -6);"#,
    ["-10"]
);

csharp_case!(
    addition_after_variable_assignment_accumulates_total,
    r#"int left = 9; int right = 11; Console.WriteLine(left + right);"#,
    ["20"]
);

csharp_case!(
    addition_in_nested_parentheses_groups_terms_first,
    r#"Console.WriteLine((3 + 4) + 5);"#,
    ["12"]
);

csharp_case!(
    addition_chain_of_three_terms_is_left_associative,
    r#"Console.WriteLine(1 + 2 + 3);"#,
    ["6"]
);

csharp_case!(
    addition_and_subtraction_share_precedence_left_to_right,
    r#"Console.WriteLine(10 + 5 - 3);"#,
    ["12"]
);

csharp_case!(
    addition_result_stored_in_new_variable,
    r#"int sum = 14 + 6; Console.WriteLine(sum);"#,
    ["20"]
);

csharp_case!(
    addition_with_unary_minus_operand,
    r#"Console.WriteLine(10 + -3);"#,
    ["7"]
);

csharp_case!(
    subtraction_yields_positive_difference_when_minuend_is_larger,
    r#"Console.WriteLine(15 - 4);"#,
    ["11"]
);

csharp_case!(
    subtraction_yields_negative_difference_when_subtrahend_is_larger,
    r#"Console.WriteLine(4 - 15);"#,
    ["-11"]
);

csharp_case!(
    subtraction_from_zero_yields_negated_subtrahend,
    r#"Console.WriteLine(0 - 8);"#,
    ["-8"]
);

csharp_case!(
    subtraction_of_zero_leaves_minuend_unchanged,
    r#"Console.WriteLine(8 - 0);"#,
    ["8"]
);

csharp_case!(
    subtraction_chain_is_left_associative,
    r#"Console.WriteLine(20 - 5 - 3);"#,
    ["12"]
);

csharp_case!(
    subtracting_negative_equivalent_to_addition,
    r#"Console.WriteLine(7 - (-3));"#,
    ["10"]
);

csharp_case!(
    subtraction_of_self_yields_zero,
    r#"int value = 42; Console.WriteLine(value - value);"#,
    ["0"]
);

csharp_case!(
    subtraction_after_addition_in_single_expression,
    r#"Console.WriteLine(2 + 8 - 5);"#,
    ["5"]
);

csharp_case!(
    subtraction_with_negative_minuend,
    r#"Console.WriteLine(-9 - 4);"#,
    ["-13"]
);

csharp_case!(
    multiplication_of_two_positive_literals,
    r#"Console.WriteLine(6 * 7);"#,
    ["42"]
);

csharp_case!(
    multiplication_by_zero_yields_zero,
    r#"Console.WriteLine(99 * 0);"#,
    ["0"]
);

csharp_case!(
    multiplication_by_one_is_identity,
    r#"Console.WriteLine(99 * 1);"#,
    ["99"]
);

csharp_case!(
    multiplication_of_negative_and_positive_yields_negative,
    r#"Console.WriteLine(-6 * 7);"#,
    ["-42"]
);

csharp_case!(
    multiplication_of_two_negatives_yields_positive,
    r#"Console.WriteLine(-6 * -7);"#,
    ["42"]
);

csharp_case!(
    multiplication_has_higher_precedence_than_addition,
    r#"Console.WriteLine(2 + 3 * 4);"#,
    ["14"]
);

csharp_case!(
    multiplication_chain_is_left_associative,
    r#"Console.WriteLine(2 * 3 * 4);"#,
    ["24"]
);

csharp_case!(
    multiplication_in_assignment_updates_product,
    r#"int factor = 5; factor = factor * 3; Console.WriteLine(factor);"#,
    ["15"]
);

csharp_case!(
    multiplication_with_unary_minus_on_operand,
    r#"Console.WriteLine(4 * -5);"#,
    ["-20"]
);

csharp_case!(
    division_truncates_positive_quotient_toward_zero,
    r#"Console.WriteLine(7 / 3);"#,
    ["2"]
);

csharp_case!(
    division_truncates_negative_dividend_toward_zero,
    r#"Console.WriteLine(-7 / 3);"#,
    ["-2"]
);

csharp_case!(
    division_with_negative_divisor_truncates_toward_zero,
    r#"Console.WriteLine(7 / -3);"#,
    ["-2"]
);

csharp_case!(
    division_of_two_negatives_truncates_toward_zero,
    r#"Console.WriteLine(-7 / -3);"#,
    ["2"]
);

csharp_case!(
    division_by_one_returns_dividend,
    r#"Console.WriteLine(42 / 1);"#,
    ["42"]
);

csharp_case!(
    division_of_zero_returns_zero,
    r#"Console.WriteLine(0 / 9);"#,
    ["0"]
);

csharp_case!(
    integer_division_discards_fraction_not_rounds_up,
    r#"Console.WriteLine(9 / 4);"#,
    ["2"]
);

csharp_case!(
    division_precedence_over_addition_in_expression,
    r#"Console.WriteLine(10 + 20 / 4);"#,
    ["15"]
);

csharp_case!(
    division_of_equal_operands_yields_one,
    r#"Console.WriteLine(15 / 15);"#,
    ["1"]
);

csharp_case!(
    division_chain_is_left_associative,
    r#"Console.WriteLine(48 / 4 / 2);"#,
    ["6"]
);

csharp_case!(
    division_with_negative_dividend_and_negative_divisor,
    r#"Console.WriteLine(-20 / -4);"#,
    ["5"]
);

csharp_case!(
    modulo_returns_remainder_for_positive_dividend,
    r#"Console.WriteLine(10 % 3);"#,
    ["1"]
);

csharp_case!(
    modulo_with_negative_dividend_keeps_dividend_sign,
    r#"Console.WriteLine(-10 % 3);"#,
    ["-1"]
);

csharp_case!(
    modulo_with_negative_divisor_keeps_dividend_sign,
    r#"Console.WriteLine(10 % -3);"#,
    ["1"]
);

csharp_case!(
    modulo_with_both_negative_keeps_dividend_sign,
    r#"Console.WriteLine(-10 % -3);"#,
    ["-1"]
);

csharp_case!(
    modulo_by_one_yields_zero,
    r#"Console.WriteLine(42 % 1);"#,
    ["0"]
);

csharp_case!(
    modulo_when_dividend_equals_divisor_yields_zero,
    r#"Console.WriteLine(7 % 7);"#,
    ["0"]
);

csharp_case!(
    modulo_when_dividend_less_than_divisor_returns_dividend,
    r#"Console.WriteLine(4 % 9);"#,
    ["4"]
);

csharp_case!(
    modulo_same_precedence_as_multiplication_left_associative,
    r#"Console.WriteLine(20 % 6 % 2);"#,
    ["0"]
);

csharp_case!(
    modulo_after_division_in_same_expression,
    r#"Console.WriteLine(17 / 5 % 3);"#,
    ["0"]
);

csharp_case!(
    modulo_with_zero_dividend_yields_zero,
    r#"Console.WriteLine(0 % 5);"#,
    ["0"]
);

csharp_case!(
    unary_minus_negates_positive_literal,
    r#"Console.WriteLine(-42);"#,
    ["-42"]
);

csharp_case!(
    unary_minus_negates_positive_variable,
    r#"int value = 42; Console.WriteLine(-value);"#,
    ["-42"]
);

csharp_case!(
    unary_minus_of_negative_yields_positive,
    r#"Console.WriteLine(-(-8));"#,
    ["8"]
);

csharp_case!(
    double_unary_minus_restores_positive_value,
    r#"int value = 15; Console.WriteLine(-(-value));"#,
    ["15"]
);

csharp_case!(
    unary_minus_binds_before_multiplication,
    r#"Console.WriteLine(-2 * 5);"#,
    ["-10"]
);

csharp_case!(
    unary_minus_binds_before_addition,
    r#"Console.WriteLine(10 + -4);"#,
    ["6"]
);

csharp_case!(
    compound_addition_assignment_accumulates_value,
    r#"int value = 10; value += 5; Console.WriteLine(value);"#,
    ["15"]
);

csharp_case!(
    compound_subtraction_assignment_reduces_value,
    r#"int value = 10; value -= 3; Console.WriteLine(value);"#,
    ["7"]
);

csharp_case!(
    compound_multiplication_assignment_scales_value,
    r#"int value = 6; value *= 4; Console.WriteLine(value);"#,
    ["24"]
);

csharp_case!(
    compound_division_assignment_truncates_quotient,
    r#"int value = 17; value /= 5; Console.WriteLine(value);"#,
    ["3"]
);

csharp_case!(
    compound_modulo_assignment_stores_remainder,
    r#"int value = 17; value %= 5; Console.WriteLine(value);"#,
    ["2"]
);

csharp_case!(
    compound_addition_with_negative_delta,
    r#"int value = 8; value += -3; Console.WriteLine(value);"#,
    ["5"]
);

csharp_case!(
    compound_subtraction_exceeding_value_flips_sign,
    r#"int value = 4; value -= 9; Console.WriteLine(value);"#,
    ["-5"]
);

csharp_case!(
    compound_multiplication_by_zero_zeros_variable,
    r#"int value = 99; value *= 0; Console.WriteLine(value);"#,
    ["0"]
);

csharp_case!(
    compound_division_on_negative_dividend,
    r#"int value = -17; value /= 5; Console.WriteLine(value);"#,
    ["-3"]
);

csharp_case!(
    compound_modulo_with_negative_dividend,
    r#"int value = -17; value %= 5; Console.WriteLine(value);"#,
    ["-2"]
);

csharp_case!(
    post_increment_returns_original_then_advances,
    r#"int value = 5; Console.WriteLine(value++); Console.WriteLine(value);"#,
    ["5", "6"]
);

csharp_case!(
    pre_increment_advances_then_returns_new_value,
    r#"int value = 5; Console.WriteLine(++value); Console.WriteLine(value);"#,
    ["6", "6"]
);

csharp_case!(
    post_decrement_returns_original_then_retreats,
    r#"int value = 5; Console.WriteLine(value--); Console.WriteLine(value);"#,
    ["5", "4"]
);

csharp_case!(
    pre_decrement_retreats_then_returns_new_value,
    r#"int value = 5; Console.WriteLine(--value); Console.WriteLine(value);"#,
    ["4", "4"]
);

csharp_case!(
    increment_on_zero_yields_one,
    r#"int value = 0; value++; Console.WriteLine(value);"#,
    ["1"]
);

csharp_case!(
    decrement_on_zero_yields_negative_one,
    r#"int value = 0; value--; Console.WriteLine(value);"#,
    ["-1"]
);

csharp_case!(
    increment_on_negative_value_moves_toward_zero,
    r#"int value = -3; value++; Console.WriteLine(value);"#,
    ["-2"]
);

csharp_case!(
    decrement_on_negative_moves_further_negative,
    r#"int value = -3; value--; Console.WriteLine(value);"#,
    ["-4"]
);

csharp_case!(
    pre_increment_used_in_addition_expression,
    r#"int value = 4; Console.WriteLine(++value + 2);"#,
    ["7"]
);

csharp_case!(
    post_increment_used_in_addition_expression,
    r#"int value = 4; Console.WriteLine(value++ + 2);"#,
    ["6"]
);

csharp_case!(
    pre_decrement_used_in_subtraction_expression,
    r#"int value = 9; int prior = value; int now = --value; Console.WriteLine(prior - now);"#,
    ["1"]
);

csharp_case!(
    chained_pre_increments_accumulate,
    r#"int left = 1; int right = 2; Console.WriteLine(++left + ++right);"#,
    ["5"]
);

csharp_case!(
    parentheses_override_multiplication_precedence,
    r#"Console.WriteLine((2 + 3) * 4);"#,
    ["20"]
);

csharp_case!(
    left_associativity_of_subtraction_without_parentheses,
    r#"Console.WriteLine(10 - 4 - 2);"#,
    ["4"]
);

csharp_case!(
    nested_parentheses_in_mixed_arithmetic_expression,
    r#"Console.WriteLine((8 - 2) * (3 + 1));"#,
    ["24"]
);

csharp_case!(
    expression_with_addition_subtraction_multiplication,
    r#"Console.WriteLine(2 + 3 * 4 - 5);"#,
    ["9"]
);

csharp_case!(
    expression_with_division_and_modulo_equal_precedence,
    r#"Console.WriteLine(20 / 3 % 2);"#,
    ["0"]
);

csharp_case!(
    expression_with_all_four_basic_operators,
    r#"Console.WriteLine(10 + 6 * 2 - 8 / 4);"#,
    ["20"]
);

csharp_case!(
    deeply_nested_parentheses_resolve_innermost_first,
    r#"Console.WriteLine(((2 + 3) * (4 - 1)) / 3);"#,
    ["5"]
);

csharp_case!(
    unary_minus_inside_parentheses_affects_grouped_sum,
    r#"Console.WriteLine(-(3 + 4));"#,
    ["-7"]
);

csharp_case!(
    mixed_expression_with_post_increment_and_addition,
    r#"int value = 3; Console.WriteLine(value++ + value);"#,
    ["7"]
);

csharp_case!(
    mixed_expression_with_pre_decrement_and_multiplication,
    r#"int value = 6; Console.WriteLine(--value * 2);"#,
    ["10"]
);

csharp_case!(
    sequential_arithmetic_statements_print_intermediate_results,
    r#"int value = 2; value = value + 3; Console.WriteLine(value); value = value * 4; Console.WriteLine(value);"#,
    ["5", "20"]
);

csharp_case!(
    reassignment_after_arithmetic_expression,
    r#"int value = 1; value = value + 2 + 3; Console.WriteLine(value);"#,
    ["6"]
);

csharp_case!(
    arithmetic_with_multiple_variables_in_one_expression,
    r#"int a = 2; int b = 3; int c = 4; Console.WriteLine(a + b * c);"#,
    ["14"]
);

csharp_case!(
    compound_assignment_chain_updates_final_value,
    r#"int value = 10; value += 5; value *= 2; value -= 8; Console.WriteLine(value);"#,
    ["22"]
);

csharp_case!(
    increment_and_compound_addition_combined,
    r#"int value = 1; value++; value += 4; Console.WriteLine(value);"#,
    ["6"]
);

csharp_case!(
    division_and_modulo_reconstruct_dividend_identity,
    r#"int dividend = 17; int divisor = 5; Console.WriteLine(dividend / divisor * divisor + dividend % divisor);"#,
    ["17"]
);

csharp_case!(
    negative_literal_in_multiplication_and_addition,
    r#"Console.WriteLine(-3 * 4 + 10);"#,
    ["-2"]
);

csharp_case!(
    subtraction_of_product_from_sum,
    r#"Console.WriteLine(20 - 3 * 5);"#,
    ["5"]
);

csharp_case!(
    modulo_precedence_with_addition,
    r#"Console.WriteLine(10 + 17 % 5);"#,
    ["12"]
);

csharp_case!(
    int_division_truncates_toward_zero_not_floor,
    r#"Console.WriteLine(-9 / 4);"#,
    ["-2"]
);

csharp_case!(
    modulo_sign_follows_dividend_not_divisor,
    r#"Console.WriteLine(-9 % 4);"#,
    ["-1"]
);

csharp_case!(
    pre_and_post_increment_on_different_variables,
    r#"int left = 2; int right = 5; Console.WriteLine(++left + right++); Console.WriteLine(left); Console.WriteLine(right);"#,
    ["8", "3", "6"]
);
