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
    addition_crosses_zero => { body: "printf(\"%d\\n\", -5 + 8);\nreturn 0;", expect: ["3"] },
    subtraction_yields_negative_result => { body: "printf(\"%d\\n\", 5 - 8);\nreturn 0;", expect: ["-3"] },
    multiplication_by_zero => { body: "printf(\"%d\\n\", 99 * 0);\nreturn 0;", expect: ["0"] },
    multiplication_of_two_negatives_is_positive => { body: "printf(\"%d\\n\", -6 * -7);\nreturn 0;", expect: ["42"] },
    multiplication_of_negative_and_positive_is_negative => { body: "printf(\"%d\\n\", -6 * 7);\nreturn 0;", expect: ["-42"] },
    integer_division_truncates_positive => { body: "printf(\"%d\\n\", 7 / 2);\nreturn 0;", expect: ["3"] },
    integer_division_truncates_toward_zero_for_negative_dividend => { body: "printf(\"%d\\n\", -7 / 2);\nreturn 0;", expect: ["-3"] },
    integer_division_truncates_toward_zero_for_negative_divisor => { body: "printf(\"%d\\n\", 7 / -2);\nreturn 0;", expect: ["-3"] },
    modulo_of_positive_numbers => { body: "printf(\"%d\\n\", 17 % 5);\nreturn 0;", expect: ["2"] },
    modulo_keeps_negative_dividend_sign => { body: "printf(\"%d\\n\", -17 % 5);\nreturn 0;", expect: ["-2"] },
    modulo_keeps_positive_dividend_sign_when_divisor_negative => { body: "printf(\"%d\\n\", 17 % -5);\nreturn 0;", expect: ["2"] },
    double_addition_preserves_fraction => { body: "printf(\"%.2f\\n\", 1.25 + 2.5);\nreturn 0;", expect: ["3.75"] },
    mixed_integer_and_double_promotes_to_double => { body: "printf(\"%.2f\\n\", 3 + 0.5);\nreturn 0;", expect: ["3.50"] },
    char_literal_promotes_to_integer => { body: "printf(\"%d\\n\", 'A' + 1);\nreturn 0;", expect: ["66"] },
    nested_parentheses_group_arithmetic => { body: "printf(\"%d\\n\", (2 + 3) * (4 + 1));\nreturn 0;", expect: ["25"] },
    chained_subtraction_is_left_associative => { body: "printf(\"%d\\n\", 20 - 5 - 3);\nreturn 0;", expect: ["12"] },
    chained_division_is_left_associative => { body: "printf(\"%d\\n\", 48 / 4 / 3);\nreturn 0;", expect: ["4"] },
    addition_with_hex_literal => { body: "printf(\"%d\\n\", 0x10 + 5);\nreturn 0;", expect: ["21"] },
    addition_with_octal_literal => { body: "printf(\"%d\\n\", 010 + 2);\nreturn 0;", expect: ["10"] },
    zero_minus_positive_number => { body: "printf(\"%d\\n\", 0 - 9);\nreturn 0;", expect: ["-9"] },
    unary_minus_applies_before_multiplication => { body: "printf(\"%d\\n\", -3 * 4);\nreturn 0;", expect: ["-12"] },
    unary_plus_leaves_value_unchanged => { body: "printf(\"%d\\n\", +7);\nreturn 0;", expect: ["7"] },
    integer_expression_can_feed_float_format => { body: "printf(\"%.1f\\n\", 5 + 2);\nreturn 0;", expect: ["7.0"] },
    floating_division_with_fractional_result => { body: "printf(\"%.2f\\n\", 7.0 / 2.0);\nreturn 0;", expect: ["3.50"] },
    integer_division_before_addition => { body: "printf(\"%d\\n\", 9 / 2 + 1);\nreturn 0;", expect: ["5"] },
    multiplication_after_parenthesized_addition => { body: "printf(\"%d\\n\", (9 / 2 + 1) * 2);\nreturn 0;", expect: ["10"] },
    double_subtraction_can_cross_zero => { body: "printf(\"%.2f\\n\", 2.5 - 4.0);\nreturn 0;", expect: ["-1.50"] },
    multiplication_with_fractional_operand => { body: "printf(\"%.2f\\n\", 6 * 0.25);\nreturn 0;", expect: ["1.50"] },
    modulo_after_multiplication_uses_product => { body: "printf(\"%d\\n\", 3 * 5 % 7);\nreturn 0;", expect: ["1"] },
    arithmetic_chain_with_negative_terms => { body: "printf(\"%d\\n\", 10 + -3 - 4);\nreturn 0;", expect: ["3"] }
}