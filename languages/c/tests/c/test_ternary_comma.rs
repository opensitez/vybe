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
    comma_operator_returns_last_expression_value => { body: "printf(\"%d\\n\", (1, 2, 3)); return 0;", expect: ["3"] },
    comma_operator_can_sequence_assignments => { body: "int x = 0; int y = 0; printf(\"%d\\n\", (x = 1, y = 2, x + y)); return 0;", expect: ["3"] },
    comma_operator_can_be_used_in_for_update_clause => { body: "for (int i = 0, j = 2; i < 2; i++, j--) printf(\"%d%d\\n\", i, j); return 0;", expect: ["02", "11"] },
    comma_operator_can_discard_intermediate_result => { body: "printf(\"%d\\n\", ((1 + 2), (3 + 4))); return 0;", expect: ["7"] },
    comma_operator_has_lower_precedence_than_assignment => { body: "int x = 0; x = (1, 2); printf(\"%d\\n\", x); return 0;", expect: ["2"] },
    ternary_can_select_integer_true_branch => { body: "printf(\"%d\\n\", 1 ? 4 : 9); return 0;", expect: ["4"] },
    ternary_can_select_integer_false_branch => { body: "printf(\"%d\\n\", 0 ? 4 : 9); return 0;", expect: ["9"] },
    ternary_can_select_string_true_branch => { body: "puts(1 ? \"yes\" : \"no\"); return 0;", expect: ["yes"] },
    ternary_can_select_string_false_branch => { body: "puts(0 ? \"yes\" : \"no\"); return 0;", expect: ["no"] },
    nested_ternary_can_pick_first_branch => { body: "int x = -1; printf(\"%d\\n\", x < 0 ? -1 : x > 0 ? 1 : 0); return 0;", expect: ["-1"] },
    nested_ternary_can_pick_last_branch => { body: "int x = 1; printf(\"%d\\n\", x < 0 ? -1 : x > 0 ? 1 : 0); return 0;", expect: ["1"] },
    ternary_condition_can_use_comparison_expression => { body: "printf(\"%d\\n\", (3 > 2) ? 8 : 9); return 0;", expect: ["8"] },
    ternary_branches_can_use_arithmetic_expressions => { body: "printf(\"%d\\n\", 1 ? 2 + 3 : 4 + 5); return 0;", expect: ["5"] },
    ternary_can_feed_character_format => { body: "printf(\"%c\\n\", 0 ? 'y' : 'n'); return 0;", expect: ["n"] },
    comma_operator_inside_ternary_condition_uses_last_value => { body: "printf(\"%d\\n\", (0, 1) ? 7 : 9); return 0;", expect: ["7"] },
    ternary_can_use_comma_operator_in_branch => { body: "printf(\"%d\\n\", 1 ? (1, 2, 3) : 0); return 0;", expect: ["3"] },
    comma_operator_can_update_variable_then_return_it => { body: "int x = 0; printf(\"%d\\n\", (x = 5, x)); return 0;", expect: ["5"] },
    ternary_expression_can_be_assigned_to_variable => { body: "int x = 0; int y = x ? 1 : 2; printf(\"%d\\n\", y); return 0;", expect: ["2"] },
    ternary_and_comma_can_combine_in_parenthesized_expression => { body: "int x = 1; printf(\"%d\\n\", (x = 0, x ? 1 : 2)); return 0;", expect: ["2"] },
    comma_operator_can_sequence_printf_side_effects => { body: "printf(\"%d\\n\", (printf(\"a\\n\"), 4)); return 0;", expect: ["a", "4"] },
    ternary_can_select_double_branch => { body: "printf(\"%.1f\\n\", 1 ? 1.5 : 2.5); return 0;", expect: ["1.5"] },
    comma_operator_can_return_character_code_expression => { body: "printf(\"%c\\n\", ('A', 'B')); return 0;", expect: ["B"] },
    ternary_false_branch_can_hold_comma_expression => { body: "printf(\"%d\\n\", 0 ? 1 : (2, 3, 4)); return 0;", expect: ["4"] },
    comma_operator_can_drive_loop_init_values => { body: "int i, j; for (i = 0, j = 3; i < 2; i++, j--) printf(\"%d%d\\n\", i, j); return 0;", expect: ["03", "12"] },
    ternary_condition_uses_nonzero_truthiness => { body: "printf(\"%d\\n\", -3 ? 1 : 0); return 0;", expect: ["1"] },
    comma_operator_result_can_feed_array_index => { body: "int arr[3] = {4, 5, 6}; printf(\"%d\\n\", arr[(0, 2)]); return 0;", expect: ["6"] },
    ternary_can_select_pointer_to_string_literal => { body: "char *text = 1 ? \"alpha\" : \"beta\"; puts(text); return 0;", expect: ["alpha"] },
    comma_operator_with_assignments_updates_both_values => { body: "int a = 0; int b = 0; (a = 2, b = 3); printf(\"%d %d\\n\", a, b); return 0;", expect: ["2 3"] },
    ternary_can_choose_comparison_result_branch => { body: "printf(\"%d\\n\", 1 ? (3 > 2) : (2 > 3)); return 0;", expect: ["1"] },
    comma_operator_can_feed_incremented_value => { body: "int x = 0; printf(\"%d\\n\", (x++, x)); return 0;", expect: ["1"] },
    ternary_false_branch_can_hold_double_expression => { body: "printf(\"%.1f\\n\", 0 ? 1.5 : 2.5); return 0;", expect: ["2.5"] },
    comma_operator_can_chain_three_assignments => { body: "int a = 0; int b = 0; int c = 0; printf(\"%d\\n\", (a = 1, b = 2, c = 3, a + b + c)); return 0;", expect: ["6"] },
    ternary_can_wrap_string_length_choice => { body: "printf(\"%d\\n\", strlen(1 ? \"cat\" : \"doge\")); return 0;", expect: ["3"] },
    comma_operator_can_feed_pointer_arithmetic_result => { body: "int arr[3] = {1, 2, 3}; int *p = arr; printf(\"%d\\n\", (p++, *p)); return 0;", expect: ["2"] },
    ternary_condition_can_use_comma_side_effect_before_test => { body: "int x = 0; printf(\"%d\\n\", (x = 1, x) ? 5 : 9); return 0;", expect: ["5"] },
    comma_operator_result_can_be_compared => { body: "printf(\"%d\\n\", ((1, 2) == 2)); return 0;", expect: ["1"] },
    ternary_can_choose_character_literal_branch => { body: "printf(\"%c\\n\", 1 ? 'A' : 'B'); return 0;", expect: ["A"] },
    comma_operator_can_sequence_function_calls_before_result => { body: "printf(\"%d\\n\", (puts(\"x\"), puts(\"y\"), 7)); return 0;", expect: ["x", "y", "7"] },
    ternary_selected_branch_can_be_pointer_difference => { body: "int arr[3] = {1, 2, 3}; printf(\"%d\\n\", 1 ? (int)(&arr[2] - &arr[0]) : 0); return 0;", expect: ["2"] },
    comma_operator_can_return_last_boolean_like_value => { body: "printf(\"%d\\n\", (0, 0, 1)); return 0;", expect: ["1"] },
    ternary_inside_comma_expression_can_use_updated_value => { body: "int x = 0; printf(\"%d\\n\", (x = 3, x > 2 ? x : 0)); return 0;", expect: ["3"] }
}
