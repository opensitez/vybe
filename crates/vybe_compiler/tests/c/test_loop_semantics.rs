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
    for_loop_init_runs_once_before_body => { body: "int x = 0; for (x = 1; x < 3; x++) printf(\"%d\\n\", x); return 0;", expect: ["1", "2"] },
    for_loop_update_runs_after_body => { body: "for (int i = 0; i < 2; printf(\"u\\n\"), i++) printf(\"b%d\\n\", i); return 0;", expect: ["b0", "u", "b1", "u"] },
    while_loop_rechecks_condition_each_iteration => { body: "int x = 0; while (x < 3) { printf(\"%d\\n\", x); x++; } return 0;", expect: ["0", "1", "2"] },
    do_while_checks_condition_after_body => { body: "int x = 3; do { printf(\"%d\\n\", x); x++; } while (x < 4); return 0;", expect: ["3"] },
    break_exits_only_current_loop_level => { body: "for (int i = 0; i < 2; i++) { for (int j = 0; j < 3; j++) { if (j == 1) break; printf(\"%d%d\\n\", i, j); } } return 0;", expect: ["00", "10"] },
    continue_skips_to_next_iteration_in_for_loop => { body: "for (int i = 0; i < 4; i++) { if (i == 2) continue; printf(\"%d\\n\", i); } return 0;", expect: ["0", "1", "3"] },
    continue_skips_to_next_iteration_in_while_loop => { body: "int i = 0; while (i < 4) { i++; if (i == 2) continue; printf(\"%d\\n\", i); } return 0;", expect: ["1", "3", "4"] },
    loop_counter_declared_in_for_is_independent_of_outer_name => { body: "int i = 10; for (int i = 0; i < 2; i++) printf(\"%d\\n\", i); printf(\"%d\\n\", i); return 0;", expect: ["0", "1", "10"] },
    loop_condition_false_initially_skips_body => { body: "for (int i = 3; i < 3; i++) puts(\"bad\"); puts(\"done\"); return 0;", expect: ["done"] },
    nested_loops_can_accumulate_pair_count => { body: "int count = 0; for (int i = 0; i < 2; i++) for (int j = 0; j < 3; j++) count++; printf(\"%d\\n\", count); return 0;", expect: ["6"] },
    while_loop_can_use_assignment_in_condition => { body: "int x = 3; while ((x = x - 1)) printf(\"%d\\n\", x); return 0;", expect: ["2", "1"] },
    do_while_can_continue_and_reach_condition_check => { body: "int x = 0; do { x++; if (x == 1) continue; printf(\"%d\\n\", x); } while (x < 3); return 0;", expect: ["2", "3"] },
    for_loop_can_have_empty_body => { body: "int i; for (i = 0; i < 3; i++); printf(\"%d\\n\", i); return 0;", expect: ["3"] },
    while_loop_can_have_single_statement_body_without_braces => { body: "int i = 0; while (i < 3) i++; printf(\"%d\\n\", i); return 0;", expect: ["3"] },
    break_inside_do_while_prevents_further_iterations => { body: "int x = 0; do { puts(\"body\"); break; } while (++x < 3); return 0;", expect: ["body"] },
    continue_in_for_loop_still_executes_update_clause => { body: "for (int i = 0; i < 3; i++) { if (i == 1) continue; printf(\"%d\\n\", i); } return 0;", expect: ["0", "2"] },
    infinite_for_loop_can_terminate_with_break => { body: "int i = 0; for (;;) { printf(\"%d\\n\", i); if (i == 1) break; i++; } return 0;", expect: ["0", "1"] },
    nested_continue_affects_only_inner_loop => { body: "for (int i = 0; i < 2; i++) { for (int j = 0; j < 3; j++) { if (j == 1) continue; printf(\"%d%d\\n\", i, j); } } return 0;", expect: ["00", "02", "10", "12"] },
    for_loop_multiple_control_variables_stay_in_sync => { body: "for (int i = 0, j = 3; i < 3; i++, j--) printf(\"%d%d\\n\", i, j); return 0;", expect: ["03", "12", "21"] },
    loop_body_can_mutate_outer_accumulator => { body: "int sum = 0; for (int i = 1; i <= 3; i++) sum += i; printf(\"%d\\n\", sum); return 0;", expect: ["6"] },
    loop_with_postfix_increment_uses_old_value_in_body_expression => { body: "for (int i = 0; i < 2;) { printf(\"%d\\n\", i++); } return 0;", expect: ["0", "1"] },
    while_loop_can_read_string_until_null => { body: "char text[] = \"go\"; int i = 0; while (text[i]) { printf(\"%c\\n\", text[i]); i++; } return 0;", expect: ["g", "o"] },
    do_while_with_false_condition_runs_once => { body: "int i = 5; do printf(\"%d\\n\", i); while (0); return 0;", expect: ["5"] },
    inner_break_does_not_skip_outer_followup_statement => { body: "for (int i = 0; i < 2; i++) { while (1) { break; } puts(\"outer\"); } return 0;", expect: ["outer", "outer"] },
    loop_can_count_down_with_decrement => { body: "for (int i = 3; i > 0; i--) printf(\"%d\\n\", i); return 0;", expect: ["3", "2", "1"] },
    while_loop_condition_using_pointer_truthiness_can_terminate => { body: "char *text = \"ok\"; int i = 0; while (text[i]) { printf(\"%c\\n\", text[i]); i++; } return 0;", expect: ["o", "k"] },
    loop_can_nest_if_else_logic => { body: "for (int i = 0; i < 3; i++) if (i % 2 == 0) puts(\"even\"); else puts(\"odd\"); return 0;", expect: ["even", "odd", "even"] },
    do_while_can_increment_then_test_limit => { body: "int i = 0; do { i++; printf(\"%d\\n\", i); } while (i < 2); return 0;", expect: ["1", "2"] },
    for_loop_can_use_comma_expression_in_condition => { body: "for (int i = 0; (i < 2, i < 3); i++) { if (i == 2) break; printf(\"%d\\n\", i); } return 0;", expect: ["0", "1"] },
    nested_loops_can_build_two_digit_grid => { body: "for (int i = 0; i < 2; i++) for (int j = 0; j < 2; j++) printf(\"%d%d\\n\", i, j); return 0;", expect: ["00", "01", "10", "11"] }
}