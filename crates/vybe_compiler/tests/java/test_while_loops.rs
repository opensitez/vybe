use crate::helpers::run_main;

#[test]
fn while_loop_zero_iterations_when_condition_false_initially() {
    let out = run_main(
        "int n = 0; while (n > 0) { System.out.println(n); } System.out.println(\"done\");",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn while_loop_counts_down_to_one() {
    let out = run_main("int n = 3; while (n > 0) { System.out.println(n); n--; }");
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn while_loop_accumulates_running_sum() {
    let out = run_main(
        "int i = 1; int sum = 0; while (i <= 4) { sum += i; i++; } System.out.println(sum);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn while_loop_with_compound_and_condition() {
    let out = run_main("int x = 0; while (x < 5 && x != 3) { System.out.println(x); x++; }");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn while_loop_single_iteration_when_bound_is_one() {
    let out = run_main("int n = 0; while (n < 1) { System.out.println(n); n++; }");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn while_loop_with_prefix_increment_in_body() {
    let out = run_main("int n = 1; while (n < 4) { System.out.println(++n); }");
    assert_eq!(out, vec!["2", "3", "4"]);
}

#[test]
fn while_break_exits_before_printing_limit_value() {
    let out =
        run_main("int n = 0; while (n < 10) { if (n == 3) break; System.out.println(n); n++; }");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn while_continue_skips_even_numbers() {
    let out = run_main(
        "int n = 0; while (n < 5) { n++; if (n % 2 == 0) continue; System.out.println(n); }",
    );
    assert_eq!(out, vec!["1", "3", "5"]);
}

#[test]
fn while_continue_skips_remaining_body_statements() {
    let out = run_main(
        "int n = 0; while (n < 3) { n++; if (n == 2) continue; System.out.println(n); System.out.println(n + 10); }",
    );
    assert_eq!(out, vec!["1", "11", "3", "13"]);
}

#[test]
fn do_while_executes_body_once_when_condition_false() {
    let out = run_main("int n = 5; do { System.out.println(n); n++; } while (n < 5);");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn do_while_repeats_until_condition_becomes_false() {
    let out = run_main("int n = 0; do { System.out.println(n); n++; } while (n < 3);");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn do_while_with_decrementing_guard_stops_at_zero() {
    let out = run_main("int n = 3; do { System.out.println(n); n--; } while (n > 0);");
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn nested_while_loops_emit_row_major_coordinates() {
    let out = run_main(
        "int r = 0; while (r < 2) { int c = 0; while (c < 2) { System.out.println(r * 10 + c); c++; } r++; }",
    );
    assert_eq!(out, vec!["0", "1", "10", "11"]);
}

#[test]
fn nested_while_inner_break_stops_inner_only() {
    let out = run_main(
        "int r = 0; while (r < 2) { int c = 0; while (c < 3) { if (c == 1) break; System.out.println(r * 10 + c); c++; } r++; }",
    );
    assert_eq!(out, vec!["0", "10"]);
}

#[test]
fn nested_while_outer_break_exits_entire_nest() {
    let out = run_main(
        "int r = 0; while (r < 3) { int c = 0; while (c < 2) { System.out.println(r); if (r == 0) { r = 5; break; } c++; } if (r == 5) break; r++; }",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn while_nested_inside_do_while_runs_both_forms() {
    let out = run_main(
        "int n = 0; do { int m = 0; while (m < 2) { System.out.println(n * 10 + m); m++; } n++; } while (n < 2);",
    );
    assert_eq!(out, vec!["0", "1", "10", "11"]);
}

#[test]
fn do_while_nested_inside_while_combines_conditions() {
    let out = run_main(
        "int i = 0; while (i < 2) { int j = 0; do { System.out.println(i + j); j++; } while (j < 2); i++; }",
    );
    assert_eq!(out, vec!["0", "1", "1", "2"]);
}

#[test]
fn three_level_nested_while_prints_depth_index() {
    let out = run_main(
        "int a = 0; while (a < 2) { int b = 0; while (b < 2) { int c = 0; while (c < 2) { System.out.println(a * 100 + b * 10 + c); c++; } b++; } a++; }",
    );
    assert_eq!(out, vec!["0", "1", "10", "11", "100", "101", "110", "111"]);
}

#[test]
fn while_with_if_filter_prints_only_above_threshold() {
    let out = run_main("int n = 0; while (n < 6) { if (n > 2) { System.out.println(n); } n++; }");
    assert_eq!(out, vec!["3", "4", "5"]);
}

#[test]
fn while_postfix_decrement_in_condition_path() {
    let out = run_main("int n = 3; while (n-- > 1) { System.out.println(n); }");
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn while_multiplies_counter_each_iteration() {
    let out = run_main("int n = 1; while (n < 20) { System.out.println(n); n *= 2; }");
    assert_eq!(out, vec!["1", "2", "4", "8", "16"]);
}

#[test]
fn do_while_continue_skips_second_print_in_body() {
    let out = run_main(
        "int n = 0; do { n++; if (n == 2) continue; System.out.println(n); } while (n < 4);",
    );
    assert_eq!(out, vec!["1", "3", "4"]);
}

#[test]
fn while_break_then_code_after_loop_still_runs() {
    let out = run_main(
        "int n = 0; while (n < 5) { if (n == 2) break; n++; } System.out.println(\"after\");",
    );
    assert_eq!(out, vec!["after"]);
}

#[test]
fn while_counts_negative_range_upward() {
    let out = run_main("int n = -2; while (n < 1) { System.out.println(n); n++; }");
    assert_eq!(out, vec!["-2", "-1", "0"]);
}

#[test]
fn while_sentinel_value_terminates_loop() {
    let out = run_main(
        "int[] data = {4, 7, -1, 9}; int i = 0; while (data[i] != -1) { System.out.println(data[i]); i++; }",
    );
    assert_eq!(out, vec!["4", "7"]);
}

#[test]
fn while_updates_outer_variable_each_pass() {
    let out = run_main(
        "int last = -1; int i = 0; while (i < 4) { last = i; i++; } System.out.println(last);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn nested_while_continue_skips_inner_column_only() {
    let out = run_main(
        "int r = 0; while (r < 2) { int c = 0; while (c < 3) { c++; if (c == 2) continue; System.out.println(r * 10 + c); } r++; }",
    );
    assert_eq!(out, vec!["1", "3", "11", "13"]);
}

#[test]
fn do_while_body_runs_before_first_condition_check() {
    let out = run_main("int n = 10; do { System.out.println(n); } while (n < 5);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn while_with_local_shadowing_does_not_leak() {
    let out = run_main(
        "int x = 0; while (x < 2) { int y = x + 5; System.out.println(y); x++; } System.out.println(x);",
    );
    assert_eq!(out, vec!["5", "6", "2"]);
}

#[test]
fn while_and_do_while_mixed_in_sequence() {
    let out = run_main(
        "int n = 0; while (n < 2) { System.out.println(\"w\" + n); n++; } int m = 0; do { System.out.println(\"d\" + m); m++; } while (m < 2);",
    );
    assert_eq!(out, vec!["w0", "w1", "d0", "d1"]);
}
