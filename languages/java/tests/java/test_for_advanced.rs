use crate::helpers::run_main;

#[test]
fn for_with_two_init_declarators_counts_jointly() {
    let out = run_main("for (int i = 0, j = 10; i < 3; i++, j--) { System.out.println(i + j); }");
    assert_eq!(out, vec!["10", "10", "10"]);
}

#[test]
fn for_with_three_init_declarators_prints_sum() {
    let out = run_main(
        "for (int a = 1, b = 2, c = 3; a < 3; a++, b++, c++) { System.out.println(a + b + c); }",
    );
    assert_eq!(out, vec!["6", "9"]);
}

#[test]
fn for_init_comma_expression_assigns_existing_variables() {
    let out = run_main(
        "int i = 0; int j = 5; for (i = 1, j = 3; i < 4; i++, j++) { System.out.println(i * j); }",
    );
    assert_eq!(out, vec!["3", "8", "15"]);
}

#[test]
fn for_update_with_two_postfix_increments() {
    let out = run_main(
        "for (int i = 0, j = 10; i < 3; i++, j--) { System.out.println(i); System.out.println(j); }",
    );
    assert_eq!(out, vec!["0", "10", "1", "9", "2", "8"]);
}

#[test]
fn for_update_with_assignment_and_increment() {
    let out = run_main("int k = 0; for (int i = 0; i < 3; i++, k += 2) { System.out.println(k); }");
    assert_eq!(out, vec!["0", "2", "4"]);
}

#[test]
fn for_update_print_side_effect_in_comma_list() {
    let out = run_main("for (int i = 0; i < 3; i++, System.out.println(i)) {}");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn for_empty_body_semicolon_only_runs_updates() {
    let out =
        run_main("int sum = 0; for (int i = 0; i < 4; i++) sum += i; System.out.println(sum);");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn for_empty_body_with_trailing_semicolon() {
    let out = run_main("for (int i = 0; i < 3; i++); System.out.println(\"done\");");
    assert_eq!(out, vec!["done"]);
}

#[test]
fn infinite_for_exits_immediately_with_break() {
    let out = run_main("for (;;) { System.out.println(1); break; }");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn infinite_for_counts_until_break() {
    let out = run_main("int n = 0; for (;;) { System.out.println(n); n++; if (n == 3) break; }");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn infinite_for_with_empty_condition_and_update() {
    let out = run_main("int i = 0; for (;; i++) { if (i == 2) break; System.out.println(i); }");
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn for_missing_init_uses_outer_counter() {
    let out = run_main("int i = 0; for (; i < 3; i++) { System.out.println(i); }");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn for_missing_update_relies_on_body_increment() {
    let out = run_main("for (int i = 0; i < 3;) { System.out.println(i); i++; }");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn labeled_break_exits_outer_for_from_inner_loop() {
    let out = run_main(
        "outer: for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (j == 1) break outer; System.out.println(i * 10 + j); } }",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn labeled_break_on_enhanced_style_nested_classic() {
    let out = run_main(
        "scan: for (int i = 0; i < 4; i++) { for (int j = 0; j < 4; j++) { if (i + j == 5) { break scan; } System.out.println(i + j); } }",
    );
    assert_eq!(
        out,
        vec!["0", "1", "2", "3", "1", "2", "3", "4", "2", "3", "4"]
    );
}

#[test]
fn labeled_continue_skips_inner_body_advances_outer() {
    let out = run_main(
        "outer: for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (j == 0) continue outer; System.out.println(i * 10 + j); } }",
    );
    assert_eq!(out, Vec::<String>::new());
}

#[test]
fn labeled_continue_jumps_to_next_outer_iteration() {
    let out = run_main(
        "loop: for (int i = 0; i < 4; i++) { if (i == 2) continue loop; System.out.println(i); }",
    );
    assert_eq!(out, vec!["0", "1", "3"]);
}

#[test]
fn nested_for_inner_break_without_label() {
    let out = run_main(
        "for (int r = 0; r < 3; r++) { for (int c = 0; c < 3; c++) { if (c == 2) break; System.out.println(r * 10 + c); } }",
    );
    assert_eq!(out, vec!["0", "1", "10", "11", "20", "21"]);
}

#[test]
fn for_with_comma_init_and_multiple_updates() {
    let out = run_main(
        "for (int x = 0, y = 9; x < 3; x++, y--) { System.out.println(x); System.out.println(y); }",
    );
    assert_eq!(out, vec!["0", "9", "1", "8", "2", "7"]);
}

#[test]
fn for_init_expression_list_before_declarations_style() {
    let out = run_main(
        "int a = 0; int b = 0; for (a = 1, b = 2; a < 4; a++, b++) { System.out.println(a + b); }",
    );
    assert_eq!(out, vec!["3", "5", "7"]);
}

#[test]
fn for_with_multiply_in_update_clause() {
    let out = run_main("for (int i = 1; i < 30; i *= 3) { System.out.println(i); }");
    assert_eq!(out, vec!["1", "3", "9", "27"]);
}

#[test]
fn for_with_subtract_in_update_clause() {
    let out = run_main("for (int i = 5; i > 0; i -= 2) { System.out.println(i); }");
    assert_eq!(out, vec!["5", "3", "1"]);
}

#[test]
fn for_body_empty_but_condition_becomes_false() {
    let out = run_main("int n = 0; for (; n < 2; n++) {} System.out.println(n);");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn infinite_for_nested_break_from_inner_if() {
    let out = run_main(
        "for (;;) { for (int j = 0; j < 5; j++) { if (j == 2) break; System.out.println(j); } break; }",
    );
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn for_with_break_in_middle_of_update_sequence() {
    let out =
        run_main("for (int i = 0; i < 10; i++) { System.out.println(i); if (i == 2) break; }");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn for_continue_skips_update_print_in_body() {
    let out =
        run_main("for (int i = 0; i < 4; i++) { if (i == 2) continue; System.out.println(i); }");
    assert_eq!(out, vec!["0", "1", "3"]);
}

#[test]
fn labeled_break_from_triple_nested_for() {
    let out = run_main(
        "done: for (int i = 0; i < 2; i++) { for (int j = 0; j < 2; j++) { for (int k = 0; k < 2; k++) { if (k == 1) break done; System.out.println(i + j + k); } } }",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn for_init_declares_final_style_counter() {
    let out = run_main("for (int i = 0; i < 2; i++) { System.out.println(i); }");
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn for_with_comma_init_zero_and_ten_offset() {
    let out = run_main(
        "for (int low = 0, high = 10; low < 3; low++, high--) { System.out.println(high - low); }",
    );
    assert_eq!(out, vec!["10", "8", "6"]);
}

#[test]
fn for_update_multiple_expressions_last_wins_visibility() {
    let out = run_main(
        "for (int i = 0; i < 3; i++, System.out.println(i * 2)) { System.out.println(i); }",
    );
    assert_eq!(out, vec!["0", "2", "1", "4", "2", "6"]);
}

#[test]
fn for_without_init_uses_prescribed_start() {
    let out = run_main("int i = 7; for (; i < 10; i++) { System.out.println(i); }");
    assert_eq!(out, vec!["7", "8", "9"]);
}

#[test]
fn for_empty_infinite_with_counter_in_body() {
    let out = run_main(
        "int steps = 0; for (;;) { steps++; if (steps == 2) break; } System.out.println(steps);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn labeled_continue_on_named_search_loop() {
    let out = run_main(
        "search: for (int i = 0; i < 5; i++) { if (i % 2 == 0) continue search; System.out.println(i); }",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn for_nested_labeled_break_after_first_match() {
    let out = run_main(
        "found: for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (i == 1 && j == 1) break found; System.out.println(i * 10 + j); } }",
    );
    assert_eq!(out, vec!["0", "1", "2", "10"]);
}

#[test]
fn for_comma_init_with_existing_vars_reset_each_entry() {
    let out = run_main(
        "int p = 0; int q = 0; for (p = 2, q = 3; p < 5; p++, q++) { System.out.println(p + q); }",
    );
    assert_eq!(out, vec!["5", "7", "9"]);
}

#[test]
fn for_with_side_effect_in_condition() {
    let out = run_main(
        "int n = 0; for (int i = 0; i < 3; i++) { while (n++ < 2) { System.out.println(n); break; } }",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn for_single_iteration_empty_body() {
    let out = run_main("for (int i = 9; i < 10; i++); System.out.println(42);");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn for_double_increment_in_update_runs_fast() {
    let out = run_main("for (int i = 0; i < 6; i += 2) { System.out.println(i); }");
    assert_eq!(out, vec!["0", "2", "4"]);
}

#[test]
fn for_labeled_continue_skips_inner_remainder() {
    let out = run_main(
        "row: for (int r = 0; r < 2; r++) { for (int c = 0; c < 3; c++) { if (c == 1) continue row; System.out.println(r * 10 + c); } }",
    );
    assert_eq!(out, vec!["0", "10"]);
}

#[test]
fn for_infinite_outer_breaks_on_flag() {
    let out = run_main(
        "boolean stop = false; for (;;) { System.out.println(1); stop = true; if (stop) break; }",
    );
    assert_eq!(out, vec!["1"]);
}
