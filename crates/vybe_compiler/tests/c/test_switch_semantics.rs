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
    switch_matches_first_case => { body: "int x = 1; switch (x) { case 1: puts(\"one\"); break; case 2: puts(\"two\"); break; } return 0;", expect: ["one"] },
    switch_matches_second_case => { body: "int x = 2; switch (x) { case 1: puts(\"one\"); break; case 2: puts(\"two\"); break; } return 0;", expect: ["two"] },
    switch_uses_default_when_no_case_matches => { body: "int x = 9; switch (x) { case 1: puts(\"one\"); break; default: puts(\"other\"); } return 0;", expect: ["other"] },
    switch_can_fall_through_two_cases => { body: "int x = 1; switch (x) { case 1: puts(\"one\"); case 2: puts(\"two\"); break; } return 0;", expect: ["one", "two"] },
    switch_case_group_can_share_one_body => { body: "int x = 2; switch (x) { case 1: case 2: puts(\"small\"); break; default: puts(\"other\"); } return 0;", expect: ["small"] },
    switch_with_char_expression_matches_char_case => { body: "char c = 'b'; switch (c) { case 'a': puts(\"a\"); break; case 'b': puts(\"b\"); break; } return 0;", expect: ["b"] },
    switch_default_can_appear_before_final_case => { body: "int x = 3; switch (x) { default: puts(\"other\"); break; case 3: puts(\"three\"); break; } return 0;", expect: ["three"] },
    switch_inside_loop_can_run_for_each_iteration => { body: "for (int i = 0; i < 3; i++) { switch (i) { case 0: puts(\"zero\"); break; case 1: puts(\"one\"); break; default: puts(\"other\"); } } return 0;", expect: ["zero", "one", "other"] },
    switch_break_does_not_break_outer_loop => { body: "for (int i = 0; i < 2; i++) { switch (i) { case 0: puts(\"zero\"); break; case 1: puts(\"one\"); break; } puts(\"loop\"); } return 0;", expect: ["zero", "loop", "one", "loop"] },
    switch_case_can_execute_expression_before_break => { body: "int x = 2; switch (x) { case 2: printf(\"%d\\n\", x + 3); break; default: puts(\"bad\"); } return 0;", expect: ["5"] },
    switch_can_use_enum_like_integer_constants => { body: "enum { START = 10, END = 20 }; int token = END; switch (token) { case START: puts(\"start\"); break; case END: puts(\"end\"); break; } return 0;", expect: ["end"] },
    switch_with_negative_case_value_matches => { body: "int x = -1; switch (x) { case -1: puts(\"neg\"); break; default: puts(\"other\"); } return 0;", expect: ["neg"] },
    switch_can_fall_through_into_default => { body: "int x = 1; switch (x) { case 1: puts(\"one\"); default: puts(\"tail\"); } return 0;", expect: ["one", "tail"] },
    switch_can_nest_inside_switch_case => { body: "int a = 1; int b = 2; switch (a) { case 1: switch (b) { case 2: puts(\"inner\"); break; default: puts(\"bad\"); } break; default: puts(\"bad\"); } return 0;", expect: ["inner"] },
    switch_default_runs_when_default_is_only_label => { body: "int x = 5; switch (x) { default: puts(\"default\"); } return 0;", expect: ["default"] },
    switch_can_match_zero_case => { body: "int x = 0; switch (x) { case 0: puts(\"zero\"); break; default: puts(\"other\"); } return 0;", expect: ["zero"] },
    switch_case_body_can_declare_local_variable => { body: "int x = 1; switch (x) { case 1: { int y = 4; printf(\"%d\\n\", y); break; } default: puts(\"other\"); } return 0;", expect: ["4"] },
    switch_can_handle_sparse_case_values => { body: "int x = 100; switch (x) { case 1: puts(\"one\"); break; case 100: puts(\"hundred\"); break; default: puts(\"other\"); } return 0;", expect: ["hundred"] },
    switch_can_use_expression_variable_as_subject => { body: "int x = 1 + 1; switch (x) { case 2: puts(\"two\"); break; default: puts(\"other\"); } return 0;", expect: ["two"] },
    switch_break_after_fallthrough_stops_later_cases => { body: "int x = 1; switch (x) { case 1: puts(\"one\"); case 2: puts(\"two\"); break; case 3: puts(\"three\"); } return 0;", expect: ["one", "two"] },
    switch_without_break_can_reach_multiple_outputs => { body: "int x = 2; switch (x) { case 1: puts(\"one\"); case 2: puts(\"two\"); case 3: puts(\"three\"); } return 0;", expect: ["two", "three"] },
    switch_with_default_between_cases_can_still_jump_to_case => { body: "int x = 3; switch (x) { case 1: puts(\"one\"); break; default: puts(\"other\"); break; case 3: puts(\"three\"); break; } return 0;", expect: ["three"] },
    switch_can_use_character_digit_case => { body: "char c = '7'; switch (c) { case '7': puts(\"digit\"); break; default: puts(\"other\"); } return 0;", expect: ["digit"] },
    switch_case_group_can_include_zero_and_one => { body: "int x = 0; switch (x) { case 0: case 1: puts(\"small\"); break; default: puts(\"other\"); } return 0;", expect: ["small"] },
    nested_switch_default_can_run_independently => { body: "int a = 1; int b = 9; switch (a) { case 1: switch (b) { default: puts(\"inner-default\"); } break; default: puts(\"outer-default\"); } return 0;", expect: ["inner-default"] },
    switch_subject_can_be_char_promoted_to_int => { body: "char c = 'A'; switch (c) { case 65: puts(\"A\"); break; default: puts(\"other\"); } return 0;", expect: ["A"] },
    switch_can_fall_through_from_zero_case => { body: "int x = 0; switch (x) { case 0: puts(\"zero\"); case 1: puts(\"one\"); break; } return 0;", expect: ["zero", "one"] },
    switch_inside_if_can_execute_selected_case => { body: "int x = 2; if (x > 0) switch (x) { case 2: puts(\"two\"); break; } return 0;", expect: ["two"] },
    switch_case_can_mutate_outer_variable => { body: "int x = 1; int y = 0; switch (x) { case 1: y = 7; break; } printf(\"%d\\n\", y); return 0;", expect: ["7"] },
    switch_default_can_follow_multiple_cases => { body: "int x = 4; switch (x) { case 1: puts(\"one\"); break; case 2: puts(\"two\"); break; default: puts(\"other\"); break; } return 0;", expect: ["other"] }
}