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
    multiplication_binds_tighter_than_addition => { declarations: "", body: "printf(\"%d\\n\", 1 + 2 * 3);\nreturn 0;", expect: ["7"] },
    parentheses_override_addition_and_multiplication => { declarations: "", body: "printf(\"%d\\n\", (1 + 2) * 3);\nreturn 0;", expect: ["9"] },
    shift_uses_result_of_addition_on_right => { declarations: "", body: "printf(\"%d\\n\", 4 << 1 + 1);\nreturn 0;", expect: ["16"] },
    shift_uses_result_of_addition_for_left_shift_amount => { declarations: "", body: "printf(\"%d\\n\", 1 << 2 + 1);\nreturn 0;", expect: ["8"] },
    division_and_multiplication_are_left_associative => { declarations: "", body: "printf(\"%d\\n\", 8 / 2 * 2);\nreturn 0;", expect: ["8"] },
    subtraction_is_left_associative => { declarations: "", body: "printf(\"%d\\n\", 8 - 3 - 2);\nreturn 0;", expect: ["3"] },
    equality_happens_after_addition => { declarations: "", body: "printf(\"%d\\n\", 1 + 2 == 4 - 1);\nreturn 0;", expect: ["1"] },
    relational_happens_before_equality => { declarations: "", body: "printf(\"%d\\n\", 1 < 2 == 1);\nreturn 0;", expect: ["1"] },
    equality_happens_before_bitwise_and => { declarations: "", body: "printf(\"%d\\n\", 3 & 1 == 1);\nreturn 0;", expect: ["1"] },
    bitwise_or_happens_before_logical_and => { declarations: "", body: "printf(\"%d\\n\", 1 && 2 | 0);\nreturn 0;", expect: ["1"] },
    logical_and_happens_before_logical_or => { declarations: "", body: "printf(\"%d\\n\", 0 || 1 && 1);\nreturn 0;", expect: ["1"] },
    ternary_uses_logical_condition_result => { declarations: "", body: "printf(\"%d\\n\", 0 || 1 ? 7 : 9);\nreturn 0;", expect: ["7"] },
    assignment_is_right_associative => { declarations: "int a = 0; int b = 0;", body: "a = b = 3;\nprintf(\"%d %d\\n\", a, b);\nreturn 0;", expect: ["3 3"] },
    unary_minus_happens_before_multiplication => { declarations: "", body: "printf(\"%d\\n\", -2 * 3 + 10);\nreturn 0;", expect: ["4"] },
    cast_happens_before_division => { declarations: "", body: "printf(\"%.2f\\n\", (double)1 / 2);\nreturn 0;", expect: ["0.50"] },
    postfix_increment_happens_after_value_use => { declarations: "int x = 3;", body: "printf(\"%d\\n\", x++ + 1);\nreturn 0;", expect: ["4"] },
    prefix_increment_happens_before_value_use => { declarations: "int x = 3;", body: "printf(\"%d\\n\", ++x + 1);\nreturn 0;", expect: ["5"] },
    array_indexing_happens_before_addition => { declarations: "int arr[3] = {5, 6, 7};", body: "printf(\"%d\\n\", arr[1] + 1);\nreturn 0;", expect: ["7"] },
    logical_not_happens_before_bitwise_or => { declarations: "", body: "printf(\"%d\\n\", !0 | 2);\nreturn 0;", expect: ["3"] },
    modulo_and_multiplication_share_left_to_right_grouping => { declarations: "", body: "printf(\"%d\\n\", 20 % 7 * 2);\nreturn 0;", expect: ["12"] },
    bitwise_shift_happens_after_additive_expression => { declarations: "", body: "printf(\"%d\\n\", 32 >> 1 + 1);\nreturn 0;", expect: ["8"] },
    ternary_binds_looser_than_addition => { declarations: "", body: "printf(\"%d\\n\", 0 ? 1 : 2 + 3);\nreturn 0;", expect: ["5"] },
    comma_operator_is_lowest_precedence => { declarations: "int x = 0;", body: "printf(\"%d\\n\", (x = 1, x + 2));\nreturn 0;", expect: ["3"] },
    bitwise_xor_happens_after_bitwise_and => { declarations: "", body: "printf(\"%d\\n\", 7 ^ 3 & 1);\nreturn 0;", expect: ["6"] },
    bitwise_or_happens_after_bitwise_xor => { declarations: "", body: "printf(\"%d\\n\", 4 | 3 ^ 1);\nreturn 0;", expect: ["6"] },
    equality_result_can_feed_ternary => { declarations: "", body: "printf(\"%d\\n\", 2 == 2 ? 9 : 1);\nreturn 0;", expect: ["9"] },
    assignment_in_condition_uses_assigned_value => { declarations: "int x = 0;", body: "if (x = 3) puts(\"true\"); else puts(\"false\");\nreturn 0;", expect: ["true"] },
    logical_or_inside_assignment_groups_first => { declarations: "int x = 0;", body: "x = 0 || 4;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["1"] },
    relational_after_shift_uses_shift_result => { declarations: "", body: "printf(\"%d\\n\", 1 << 2 < 5);\nreturn 0;", expect: ["1"] },
    additive_after_multiplicative_and_shift => { declarations: "", body: "printf(\"%d\\n\", 1 + 2 << 2);\nreturn 0;", expect: ["12"] }
}
