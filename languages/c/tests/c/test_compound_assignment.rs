use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    plus_equals_adds_right_operand => { declarations: "int x = 5;", body: "x += 7;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["12"] },
    minus_equals_subtracts_right_operand => { declarations: "int x = 5;", body: "x -= 7;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["-2"] },
    times_equals_multiplies_existing_value => { declarations: "int x = 5;", body: "x *= 7;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["35"] },
    divide_equals_uses_integer_division => { declarations: "int x = 17;", body: "x /= 5;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["3"] },
    modulo_equals_stores_remainder => { declarations: "int x = 17;", body: "x %= 5;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["2"] },
    bitwise_and_equals_masks_bits => { declarations: "int x = 14;", body: "x &= 10;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["10"] },
    bitwise_or_equals_sets_bits => { declarations: "int x = 8;", body: "x |= 3;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["11"] },
    bitwise_xor_equals_toggles_bits => { declarations: "int x = 14;", body: "x ^= 3;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["13"] },
    left_shift_equals_shifts_bits_left => { declarations: "int x = 3;", body: "x <<= 2;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["12"] },
    right_shift_equals_shifts_bits_right => { declarations: "int x = 16;", body: "x >>= 3;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["2"] },
    plus_equals_on_array_element_updates_slot => { declarations: "int arr[2] = {1, 2};", body: "arr[1] += 5;\nprintf(\"%d\\n\", arr[1]);\nreturn 0;", expect: ["7"] },
    minus_equals_on_array_element_updates_slot => { declarations: "int arr[2] = {1, 9};", body: "arr[1] -= 4;\nprintf(\"%d\\n\", arr[1]);\nreturn 0;", expect: ["5"] },
    times_equals_on_expression_result_uses_old_value_once => { declarations: "int x = 4;", body: "x *= x + 1;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["20"] },
    divide_equals_on_negative_value_truncates_toward_zero => { declarations: "int x = -9;", body: "x /= 2;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["-4"] },
    modulo_equals_keeps_dividend_sign => { declarations: "int x = -9;", body: "x %= 2;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["-1"] },
    compound_assignments_can_chain_over_statements => { declarations: "int x = 2;", body: "x += 3;\nx *= 4;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["20"] },
    plus_equals_returns_assigned_value_in_expression => { declarations: "int x = 1; int y = 0;", body: "y = (x += 4);\nprintf(\"%d %d\\n\", x, y);\nreturn 0;", expect: ["5 5"] },
    shift_equals_combines_with_additive_rhs => { declarations: "int x = 1;", body: "x <<= 1 + 2;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["8"] },
    xor_equals_can_zero_same_bits => { declarations: "int x = 7;", body: "x ^= 7;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["0"] },
    and_equals_can_zero_all_bits => { declarations: "int x = 7;", body: "x &= 0;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["0"] },
    or_equals_can_leave_value_unchanged => { declarations: "int x = 7;", body: "x |= 0;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["7"] },
    compound_assignment_with_parenthesized_lhs_expression => { declarations: "int arr[2] = {2, 3}; int i = 0;", body: "(arr[i]) += 5;\nprintf(\"%d\\n\", arr[0]);\nreturn 0;", expect: ["7"] },
    compound_assignment_on_double_keeps_fraction => { declarations: "double x = 1.5;", body: "x += 0.25;\nprintf(\"%.2f\\n\", x);\nreturn 0;", expect: ["1.75"] },
    compound_assignment_on_char_promotes_and_stores_result => { declarations: "char c = 'A';", body: "c += 2;\nprintf(\"%c\\n\", c);\nreturn 0;", expect: ["C"] },
    right_shift_equals_on_even_number_halves_repeatedly => { declarations: "int x = 64;", body: "x >>= 1;\nx >>= 1;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["16"] }
}
