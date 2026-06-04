use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    int_plus_double_promotes_to_double => { body: "printf(\"%.1f\\n\", 2 + 0.5); return 0;", expect: ["2.5"] },
    char_plus_int_promotes_char_to_int => { body: "printf(\"%d\\n\", 'A' + 2); return 0;", expect: ["67"] },
    char_minus_char_produces_integer_difference => { body: "printf(\"%d\\n\", 'd' - 'a'); return 0;", expect: ["3"] },
    unsigned_and_signed_expression_can_be_formatted_unsigned => { body: "printf(\"%u\\n\", 1u + 2); return 0;", expect: ["3"] },
    float_literal_in_expression_promotes_integer_side => { body: "printf(\"%.2f\\n\", 3 * 0.5); return 0;", expect: ["1.50"] },
    division_by_double_promotes_integer_dividend => { body: "printf(\"%.2f\\n\", 3 / 2.0); return 0;", expect: ["1.50"] },
    comparison_between_int_and_double_uses_numeric_conversion => { body: "printf(\"%d\\n\", 3 == 3.0); return 0;", expect: ["1"] },
    ternary_between_int_and_double_has_double_common_type => { body: "printf(\"%.1f\\n\", 1 ? 3 : 4.5); return 0;", expect: ["3.0"] },
    unary_plus_on_char_still_promotes_to_int => { body: "printf(\"%d\\n\", +'A'); return 0;", expect: ["65"] },
    bitwise_ops_on_char_literals_use_integer_promotions => { body: "printf(\"%d\\n\", ('A' | 32)); return 0;", expect: ["97"] },
    shift_on_char_literal_uses_promoted_integer => { body: "printf(\"%d\\n\", ('A' >> 1)); return 0;", expect: ["32"] },
    multiplication_of_int_and_double_returns_double => { body: "printf(\"%.1f\\n\", 4 * 2.5); return 0;", expect: ["10.0"] },
    subtraction_of_double_and_int_returns_double => { body: "printf(\"%.1f\\n\", 5.5 - 2); return 0;", expect: ["3.5"] },
    modulo_keeps_integer_type_after_char_promotion => { body: "printf(\"%d\\n\", 'd' % 10); return 0;", expect: ["0"] },
    prefix_increment_on_char_promotes_for_formatting => { body: "char c = 'a'; printf(\"%d\\n\", ++c); return 0;", expect: ["98"] },
    array_of_char_values_promotes_element_in_arithmetic => { body: "char letters[2] = {'a', 'b'}; printf(\"%d\\n\", letters[0] + 1); return 0;", expect: ["98"] },
    signed_zero_comparison_with_double_zero_is_true => { body: "printf(\"%d\\n\", 0 == 0.0); return 0;", expect: ["1"] },
    negative_int_plus_double_promotes_to_double => { body: "printf(\"%.1f\\n\", -2 + 0.5); return 0;", expect: ["-1.5"] },
    conditional_operator_with_char_and_int_uses_int_result => { body: "printf(\"%d\\n\", 1 ? 'A' : 3); return 0;", expect: ["65"] },
    comparison_after_promotion_of_char_and_int_is_true => { body: "printf(\"%d\\n\", 'A' < 100); return 0;", expect: ["1"] },
    double_assigned_from_integer_division_after_cast_preserves_fraction => { body: "double value = (double)3 / 2; printf(\"%.2f\\n\", value); return 0;", expect: ["1.50"] },
    integer_result_from_mixed_expression_can_be_cast_back => { body: "printf(\"%d\\n\", (int)(2 + 0.75)); return 0;", expect: ["2"] },
    long_and_int_addition_can_be_printed_as_long => { body: "printf(\"%ld\\n\", 2l + 3); return 0;", expect: ["5"] },
    unsigned_char_promotion_can_print_255 => { body: "unsigned char c = 255; printf(\"%u\\n\", c); return 0;", expect: ["255"] },
    mixed_comparison_between_unsigned_and_int_positive_values_is_true => { body: "unsigned int u = 5; int i = 5; printf(\"%d\\n\", u == i); return 0;", expect: ["1"] },
    double_result_from_ternary_can_feed_float_format => { body: "printf(\"%.1f\\n\", 0 ? 1.5 : 2.5); return 0;", expect: ["2.5"] },
    arithmetic_on_boolean_like_results_uses_ints => { body: "printf(\"%d\\n\", (3 > 2) + (2 > 3)); return 0;", expect: ["1"] },
    integer_promotion_applies_before_bitwise_not_on_char => { body: "signed char c = 0; printf(\"%d\\n\", ~c); return 0;", expect: ["-1"] },
    promotion_of_char_in_function_argument_matches_int_parameter => { body: "int take_int(int x) { return x; } printf(\"%d\\n\", take_int('A')); return 0;", expect: ["65"] },
    promotion_in_comma_expression_uses_last_operand_type => { body: "printf(\"%.1f\\n\", (1, 2.5)); return 0;", expect: ["2.5"] }
}
