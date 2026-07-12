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
    else_if_chain_selects_middle_branch => { body: "int x = 5;\nif (x < 0) puts(\"neg\"); else if (x == 5) puts(\"five\"); else puts(\"other\");\nreturn 0;", expect: ["five"] },
    else_if_chain_selects_final_else_branch => { body: "int x = 9;\nif (x < 0) puts(\"neg\"); else if (x == 5) puts(\"five\"); else puts(\"other\");\nreturn 0;", expect: ["other"] },
    dangling_else_binds_to_nearest_if => { body: "int a = 1; int b = 0;\nif (a) if (b) puts(\"inner\"); else puts(\"else\");\nreturn 0;", expect: ["else"] },
    if_without_braces_controls_single_statement_only => { body: "int x = 1;\nif (x) puts(\"one\"); puts(\"two\");\nreturn 0;", expect: ["one", "two"] },
    nested_if_else_can_pick_deep_branch => { body: "int x = 4;\nif (x > 0) { if (x % 2 == 0) puts(\"even\"); else puts(\"odd\"); } else puts(\"neg\");\nreturn 0;", expect: ["even"] },
    ternary_operator_can_choose_string_result => { body: "int x = 4;\nputs(x % 2 == 0 ? \"even\" : \"odd\");\nreturn 0;", expect: ["even"] },
    nested_ternary_picks_middle_value => { body: "int x = 0;\nprintf(\"%d\\n\", x < 0 ? -1 : x > 0 ? 1 : 0);\nreturn 0;", expect: ["0"] },
    ternary_can_return_double_expression => { body: "int x = 1;\nprintf(\"%.1f\\n\", x ? 1.5 : 2.5);\nreturn 0;", expect: ["1.5"] },
    ternary_condition_uses_nonzero_truthiness => { body: "int x = -3;\nprintf(\"%d\\n\", x ? 7 : 9);\nreturn 0;", expect: ["7"] },
    ternary_can_be_nested_inside_printf_arguments => { body: "int x = 2;\nprintf(\"%s\\n\", x == 2 ? \"two\" : \"other\");\nreturn 0;", expect: ["two"] },
    if_condition_can_use_assignment_expression => { body: "int x = 0;\nif (x = 3) puts(\"true\"); else puts(\"false\");\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["true", "3"] },
    nested_if_without_else_can_skip_all_output_until_afterwards => { body: "int x = 0;\nif (x) if (1) puts(\"bad\");\nputs(\"done\");\nreturn 0;", expect: ["done"] },
    if_else_can_compare_strings_via_strcmp_result => { body: "char *a = \"cat\"; char *b = \"cat\";\nif (a == b) puts(\"same\"); else puts(\"diff\");\nreturn 0;", expect: ["same"] },
    ternary_branches_can_be_parenthesized_expressions => { body: "int x = 1;\nprintf(\"%d\\n\", x ? (2 + 3) : (4 + 5));\nreturn 0;", expect: ["5"] },
    ternary_can_feed_assignment => { body: "int x = 0; int y = x ? 1 : 2;\nprintf(\"%d\\n\", y);\nreturn 0;", expect: ["2"] },
    nested_else_if_chain_can_pick_first_branch => { body: "int x = -1;\nif (x < 0) puts(\"neg\"); else if (x == 0) puts(\"zero\"); else puts(\"pos\");\nreturn 0;", expect: ["neg"] },
    conditional_expression_can_select_char_literal => { body: "int x = 1;\nprintf(\"%c\\n\", x ? 'y' : 'n');\nreturn 0;", expect: ["y"] },
    if_else_blocks_can_have_independent_scopes => { body: "int x = 1;\nif (x) { int y = 2; printf(\"%d\\n\", y); } else { int y = 3; printf(\"%d\\n\", y); }\nreturn 0;", expect: ["2"] },
    ternary_can_select_pointer_to_string => { body: "int x = 0; char *text = x ? \"yes\" : \"no\";\nputs(text);\nreturn 0;", expect: ["no"] },
    if_inside_else_branch_can_still_match => { body: "int x = 0;\nif (x) puts(\"if\"); else if (!x) puts(\"else-if\");\nreturn 0;", expect: ["else-if"] }
}
