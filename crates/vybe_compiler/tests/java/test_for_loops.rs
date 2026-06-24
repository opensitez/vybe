use crate::helpers::run_main;

#[test]
fn classic_for_counts_from_zero_to_exclusive_bound() {
    let out = run_main("for (int i = 0; i < 3; i++) { System.out.println(i); }");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn classic_for_with_nonzero_start_index() {
    let out = run_main("for (int i = 5; i < 8; i++) { System.out.println(i); }");
    assert_eq!(out, vec!["5", "6", "7"]);
}

#[test]
fn classic_for_with_step_greater_than_one() {
    let out = run_main("for (int i = 0; i < 10; i += 2) { System.out.println(i); }");
    assert_eq!(out, vec!["0", "2", "4", "6", "8"]);
}

#[test]
fn classic_for_counts_down_with_decrement() {
    let out = run_main("for (int i = 3; i > 0; i--) { System.out.println(i); }");
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn classic_for_zero_iterations_when_start_equals_bound() {
    let out = run_main(
        "for (int i = 5; i < 5; i++) { System.out.println(i); } System.out.println(\"done\");",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn classic_for_never_enters_when_start_above_bound() {
    let out = run_main(
        "for (int i = 10; i < 3; i++) { System.out.println(i); } System.out.println(\"skip\");",
    );
    assert_eq!(out, vec!["skip"]);
}

#[test]
fn classic_for_single_iteration() {
    let out = run_main("for (int i = 0; i < 1; i++) { System.out.println(i); }");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn classic_for_accumulates_sum() {
    let out = run_main(
        "int sum = 0; for (int i = 1; i <= 4; i++) { sum += i; } System.out.println(sum);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn classic_for_body_with_multiple_statements() {
    let out = run_main(
        "for (int i = 0; i < 2; i++) { System.out.println(i); System.out.println(i * 10); }",
    );
    assert_eq!(out, vec!["0", "0", "1", "10"]);
}

#[test]
fn classic_for_update_uses_postfix_increment() {
    let out = run_main(
        "int last = -1; for (int i = 0; i < 3; i++) { last = i; } System.out.println(last);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn classic_for_condition_uses_comparison_expression() {
    let out = run_main("for (int i = 0; i * i < 20; i++) { System.out.println(i); }");
    assert_eq!(out, vec!["0", "1", "2", "3", "4"]);
}

#[test]
fn loop_variable_not_visible_after_for_loop() {
    let out = run_main(
        "for (int i = 0; i < 2; i++) {} int x = 7; System.out.println(x);",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn inner_for_shadows_outer_loop_variable() {
    let out = run_main(
        "for (int i = 0; i < 2; i++) { for (int i = 0; i < 2; i++) { System.out.println(i); } System.out.println(\"row\"); }",
    );
    assert_eq!(out, vec!["0", "1", "row", "0", "1", "row"]);
}

#[test]
fn enhanced_for_iterates_int_array_elements() {
    let out = run_main(
        "int[] nums = {10, 20, 30}; for (int v : nums) { System.out.println(v); }",
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn enhanced_for_on_single_element_array() {
    let out = run_main("int[] nums = {42}; for (int v : nums) { System.out.println(v); }");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn enhanced_for_accumulates_array_total() {
    let out = run_main(
        "int[] nums = {1, 2, 3, 4}; int total = 0; for (int v : nums) { total += v; } System.out.println(total);",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn enhanced_for_with_string_array() {
    let out = run_main(
        r#"String[] words = {"a", "b", "c"}; for (String w : words) { System.out.println(w); }"#,
    );
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn enhanced_for_empty_array_runs_zero_times() {
    let out = run_main(
        "int[] nums = new int[0]; for (int v : nums) { System.out.println(v); } System.out.println(\"end\");",
    );
    assert_eq!(out, vec!["end"]);
}

#[test]
fn nested_classic_for_produces_cartesian_pairs() {
    let out = run_main(
        "for (int r = 0; r < 2; r++) { for (int c = 0; c < 2; c++) { System.out.println(r * 10 + c); } }",
    );
    assert_eq!(out, vec!["0", "1", "10", "11"]);
}

#[test]
fn nested_for_three_levels() {
    let out = run_main(
        "for (int i = 0; i < 2; i++) { for (int j = 0; j < 2; j++) { for (int k = 0; k < 2; k++) { System.out.println(i + j + k); } } }",
    );
    assert_eq!(out, vec!["0", "1", "1", "2", "1", "2", "2", "3"]);
}

#[test]
fn break_exits_classic_for_early() {
    let out = run_main(
        "for (int i = 0; i < 10; i++) { if (i == 3) break; System.out.println(i); }",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn break_on_first_match_in_enhanced_for() {
    let out = run_main(
        "int[] nums = {5, 10, 15, 20}; for (int v : nums) { if (v == 15) break; System.out.println(v); }",
    );
    assert_eq!(out, vec!["5", "10"]);
}

#[test]
fn break_in_inner_loop_does_not_exit_outer() {
    let out = run_main(
        "for (int r = 0; r < 3; r++) { for (int c = 0; c < 3; c++) { if (c == 1) break; System.out.println(r * 10 + c); } }",
    );
    assert_eq!(out, vec!["0", "10", "20"]);
}

#[test]
fn continue_skips_even_iterations() {
    let out = run_main(
        "for (int i = 0; i < 5; i++) { if (i % 2 == 0) continue; System.out.println(i); }",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn continue_in_enhanced_for_skips_selected_values() {
    let out = run_main(
        "int[] nums = {1, 2, 3, 4, 5}; for (int v : nums) { if (v == 3) continue; System.out.println(v); }",
    );
    assert_eq!(out, vec!["1", "2", "4", "5"]);
}

#[test]
fn continue_skips_remaining_body_statements() {
    let out = run_main(
        "for (int i = 0; i < 3; i++) { if (i == 1) continue; System.out.println(i); System.out.println(i + 100); }",
    );
    assert_eq!(out, vec!["0", "100", "2", "102"]);
}

#[test]
fn for_with_if_inside_filters_output() {
    let out = run_main(
        "for (int i = 0; i < 6; i++) { if (i > 2) { System.out.println(i); } }",
    );
    assert_eq!(out, vec!["3", "4", "5"]);
}

#[test]
fn classic_for_reassigns_outer_variable_each_iteration() {
    let out = run_main(
        "int last = -1; for (int i = 0; i < 4; i++) { last = i; } System.out.println(last);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn nested_enhanced_for_over_rows() {
    let out = run_main(
        "int[][] grid = {{1, 2}, {3, 4}}; for (int[] row : grid) { for (int v : row) { System.out.println(v); } }",
    );
    assert_eq!(out, vec!["1", "2", "3", "4"]);
}

#[test]
fn classic_for_with_negative_start_counts_up() {
    let out = run_main("for (int i = -2; i < 1; i++) { System.out.println(i); }");
    assert_eq!(out, vec!["-2", "-1", "0"]);
}

#[test]
fn break_after_printing_last_allowed_value() {
    let out = run_main(
        "for (int i = 0; i < 100; i++) { System.out.println(i); if (i == 2) break; }",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn for_loop_declares_separate_scope_per_nesting_level() {
    let out = run_main(
        "int x = 0; for (int i = 0; i < 2; i++) { int y = i + 1; x += y; } System.out.println(x);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn enhanced_for_reads_but_outer_counter_tracks_length() {
    let out = run_main(
        "int[] nums = {7, 8, 9}; int count = 0; for (int v : nums) { count++; } System.out.println(count);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn classic_for_with_multiply_in_update() {
    let out = run_main("for (int i = 1; i < 20; i *= 2) { System.out.println(i); }");
    assert_eq!(out, vec!["1", "2", "4", "8", "16"]);
}

#[test]
fn nested_break_and_continue_together() {
    let out = run_main(
        "for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (j == 0) continue; if (j == 2) break; System.out.println(i * 10 + j); } }",
    );
    assert_eq!(out, vec!["1", "11", "21"]);
}
