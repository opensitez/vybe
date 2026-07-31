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

// ── `/` on integral VARIABLES, across C's type spellings ───────────────────
//
// Every division case above divides two LITERALS, and a literal short-circuits
// the integral check before any type hint is read. So the spellings themselves
// were untested — including the multi-word ones (`unsigned int`, `long long`)
// that only a substring match resolves, and `char`, which is an integer type in
// C however unlike one it looks.
//
// Expected values measured by compiling and running the same program with the
// system `cc`, not derived from the standard.
c_cases! {
    integer_division_on_int_variables_truncates => {
        body: "int a = 7, b = 2;\nprintf(\"%d\\n\", a / b);\nreturn 0;",
        expect: ["3"]
    },
    integer_division_on_long_variables_truncates => {
        body: "long a = 17, b = 5;\nprintf(\"%ld\\n\", a / b);\nreturn 0;",
        expect: ["3"]
    },
    integer_division_on_unsigned_int_variables_truncates => {
        body: "unsigned int a = 9, b = 4;\nprintf(\"%u\\n\", a / b);\nreturn 0;",
        expect: ["2"]
    },
    integer_division_on_long_long_variables_truncates => {
        body: "long long a = 22, b = 7;\nprintf(\"%lld\\n\", a / b);\nreturn 0;",
        expect: ["3"]
    },
    integer_division_on_short_variables_truncates => {
        body: "short a = 9, b = 2;\nprintf(\"%d\\n\", a / b);\nreturn 0;",
        expect: ["4"]
    },
    // `char` is an integer type in C: 7 / 2 is 3, not 3.5.
    integer_division_on_char_variables_truncates => {
        body: "char a = 7, b = 2;\nprintf(\"%d\\n\", a / b);\nreturn 0;",
        expect: ["3"]
    },
    // The other side of the gate — widening the integral spellings must not
    // make real division truncate.
    division_on_double_variables_does_not_truncate => {
        body: "double a = 7.0, b = 2.0;\nprintf(\"%.1f\\n\", a / b);\nreturn 0;",
        expect: ["3.5"]
    },
}
