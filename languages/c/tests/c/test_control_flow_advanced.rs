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
    switch_can_fall_through_without_break => { body: "int x = 1;\nswitch (x) { case 1: puts(\"one\"); case 2: puts(\"two\"); break; default: puts(\"other\"); }\nreturn 0;", expect: ["one", "two"] },
    switch_default_can_appear_before_later_case_labels => { body: "int x = 2;\nswitch (x) { default: puts(\"default\"); break; case 2: puts(\"two\"); break; }\nreturn 0;", expect: ["two"] },
    do_while_executes_body_once_when_condition_false => { body: "int x = 0;\ndo { puts(\"once\"); } while (x);\nreturn 0;", expect: ["once"] },
    while_loop_break_exits_early => { body: "int i = 0;\nwhile (1) { if (i == 3) break; printf(\"%d\\n\", i); i++; }\nreturn 0;", expect: ["0", "1", "2"] },
    while_loop_continue_skips_body_tail => { body: "int i = 0;\nwhile (i < 4) { i++; if (i == 2) continue; printf(\"%d\\n\", i); }\nreturn 0;", expect: ["1", "3", "4"] },
    for_loop_without_init_can_use_external_counter => { body: "int i = 0;\nfor (; i < 3; i++) printf(\"%d\\n\", i);\nreturn 0;", expect: ["0", "1", "2"] },
    for_loop_without_update_can_increment_inside_body => { body: "for (int i = 0; i < 3; ) { printf(\"%d\\n\", i); i++; }\nreturn 0;", expect: ["0", "1", "2"] },
    for_loop_without_condition_can_break_manually => { body: "for (int i = 0; ; i++) { if (i == 3) break; printf(\"%d\\n\", i); }\nreturn 0;", expect: ["0", "1", "2"] },
    nested_break_only_exits_inner_loop => { body: "for (int i = 0; i < 2; i++) { for (int j = 0; j < 3; j++) { if (j == 1) break; printf(\"%d%d\\n\", i, j); } }\nreturn 0;", expect: ["00", "10"] },
    nested_continue_only_skips_inner_iteration => { body: "for (int i = 0; i < 2; i++) { for (int j = 0; j < 3; j++) { if (j == 1) continue; printf(\"%d%d\\n\", i, j); } }\nreturn 0;", expect: ["00", "02", "10", "12"] },
    switch_on_character_constant_matches_case => { body: "char c = 'b';\nswitch (c) { case 'a': puts(\"a\"); break; case 'b': puts(\"b\"); break; default: puts(\"other\"); }\nreturn 0;", expect: ["b"] },
    switch_inside_loop_can_vary_by_iteration => { body: "for (int i = 0; i < 3; i++) { switch (i) { case 0: puts(\"zero\"); break; case 1: puts(\"one\"); break; default: puts(\"other\"); break; } }\nreturn 0;", expect: ["zero", "one", "other"] },
    empty_for_body_can_still_update_counter => { body: "int i;\nfor (i = 0; i < 3; i++) ;\nprintf(\"%d\\n\", i);\nreturn 0;", expect: ["3"] },
    empty_while_body_can_still_change_counter_in_condition_block => { body: "int i = 0;\nwhile (i < 3) i++;\nprintf(\"%d\\n\", i);\nreturn 0;", expect: ["3"] },
    break_from_switch_does_not_break_outer_loop => { body: "for (int i = 0; i < 2; i++) { switch (i) { case 0: puts(\"zero\"); break; case 1: puts(\"one\"); break; } puts(\"loop\"); }\nreturn 0;", expect: ["zero", "loop", "one", "loop"] },
    continue_in_for_loop_skips_post_continue_statements_only => { body: "for (int i = 0; i < 3; i++) { if (i == 1) continue; printf(\"%d\\n\", i); puts(\"tail\"); }\nreturn 0;", expect: ["0", "tail", "2", "tail"] },
    nested_switch_can_match_inner_default => { body: "int x = 1; int y = 3;\nswitch (x) { case 1: switch (y) { case 2: puts(\"two\"); break; default: puts(\"inner\"); break; } break; default: puts(\"outer\"); }\nreturn 0;", expect: ["inner"] },
    while_loop_condition_can_use_assignment => { body: "int x = 3;\nwhile ((x = x - 1)) printf(\"%d\\n\", x);\nreturn 0;", expect: ["2", "1"] },
    do_while_can_continue_and_recheck_condition => { body: "int x = 0;\ndo { x++; if (x == 1) continue; printf(\"%d\\n\", x); } while (x < 3);\nreturn 0;", expect: ["2", "3"] },
    for_loop_init_declares_scope_local_variable => { body: "for (int i = 0; i < 2; i++) printf(\"%d\\n\", i);\nreturn 0;", expect: ["0", "1"] },
    switch_case_can_share_label_body => { body: "int x = 2;\nswitch (x) { case 1: case 2: puts(\"small\"); break; default: puts(\"other\"); }\nreturn 0;", expect: ["small"] },
    for_loop_multiple_clauses_can_update_two_variables => { body: "for (int i = 0, j = 3; i < 3; i++, j--) printf(\"%d%d\\n\", i, j);\nreturn 0;", expect: ["03", "12", "21"] },
    while_loop_can_nest_if_else_logic => { body: "int i = 0;\nwhile (i < 3) { if (i % 2 == 0) puts(\"even\"); else puts(\"odd\"); i++; }\nreturn 0;", expect: ["even", "odd", "even"] },
    do_while_with_break_can_exit_before_condition_check => { body: "int x = 0;\ndo { puts(\"body\"); break; } while (x < 10);\nreturn 0;", expect: ["body"] },
    switch_default_runs_when_no_case_matches => { body: "int x = 99;\nswitch (x) { case 1: puts(\"one\"); break; default: puts(\"default\"); }\nreturn 0;", expect: ["default"] },
    nested_loops_can_use_break_and_continue_together => { body: "for (int i = 0; i < 2; i++) { for (int j = 0; j < 3; j++) { if (j == 1) continue; if (j == 2) break; printf(\"%d%d\\n\", i, j); } }\nreturn 0;", expect: ["00", "10"] },
    infinite_loop_with_break_can_terminate_normally => { body: "int x = 0;\nfor (;;) { printf(\"%d\\n\", x); if (x == 1) break; x++; }\nreturn 0;", expect: ["0", "1"] },
    switch_can_use_expression_in_case_body => { body: "int x = 2;\nswitch (x) { case 2: printf(\"%d\\n\", x + 3); break; default: puts(\"bad\"); }\nreturn 0;", expect: ["5"] },
    while_loop_can_use_pointer_like_truth_value => { body: "char *text = \"go\"; int i = 0;\nwhile (text[i]) { printf(\"%c\\n\", text[i]); i++; }\nreturn 0;", expect: ["g", "o"] },
    for_loop_condition_false_initially_skips_body => { body: "for (int i = 3; i < 3; i++) puts(\"bad\");\nputs(\"done\");\nreturn 0;", expect: ["done"] }
}
