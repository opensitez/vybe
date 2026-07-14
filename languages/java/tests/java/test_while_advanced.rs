use crate::helpers::run_main;

#[test]
fn while_complex_condition_with_and_or_and_not() {
    let out = run_main(
        "int x = 4; int y = 7; while (x < 10 && (y > 5 || x == 4) && !(x < 0)) { System.out.println(x); x++; }",
    );
    assert_eq!(out, vec!["4", "5", "6", "7", "8", "9"]);
}

#[test]
fn while_condition_with_nested_parentheses_and_modulo() {
    let out =
        run_main("int n = 1; while ((n % 3 != 0) && (n < 5)) { System.out.println(n); n++; }");
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn while_sentinel_reads_until_negative_one() {
    let out = run_main(
        "int[] data = {4, 7, 12, -1, 99}; int i = 0; while (data[i] != -1) { System.out.println(data[i]); i++; }",
    );
    assert_eq!(out, vec!["4", "7", "12"]);
}

#[test]
fn while_sentinel_on_string_array_until_null_marker() {
    let out = run_main(
        r#"String[] words = {"a", "b", "done", "z"}; int i = 0; while (!words[i].equals("done")) { System.out.println(words[i]); i++; }"#,
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn while_assignment_in_condition_reads_array() {
    let out = run_main(
        "int[] vals = {3, 5, 0}; int i = 0; int v; while ((v = vals[i++]) != 0) { System.out.println(v); }",
    );
    assert_eq!(out, vec!["3", "5"]);
}

#[test]
fn do_while_menu_simulation_three_choices() {
    let out = run_main(
        "int choice = 0; int round = 0; do { round++; choice = round; System.out.println(\"menu\" + choice); } while (choice < 3);",
    );
    assert_eq!(out, vec!["menu1", "menu2", "menu3"]);
}

#[test]
fn do_while_menu_exit_on_quit_option() {
    let out = run_main(
        "int opt = 0; do { opt++; System.out.println(opt); if (opt == 2) break; } while (opt < 5);",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn do_while_menu_repeats_until_valid_input() {
    let out = run_main(
        "int input = -1; int tries = 0; do { tries++; input = tries - 1; System.out.println(input); } while (input < 2);",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn nested_while_inner_break_leaves_outer_running() {
    let out = run_main(
        "int r = 0; while (r < 2) { int c = 0; while (c < 4) { if (c == 2) break; System.out.println(r * 10 + c); c++; } r++; }",
    );
    assert_eq!(out, vec!["0", "1", "10", "11"]);
}

#[test]
fn nested_while_outer_break_via_flag() {
    let out = run_main(
        "int r = 0; boolean stop = false; while (r < 5 && !stop) { int c = 0; while (c < 3) { System.out.println(r); if (r == 1) { stop = true; break; } c++; } r++; }",
    );
    assert_eq!(out, vec!["0", "0", "0", "1"]);
}

#[test]
fn nested_while_triple_break_from_innermost() {
    let out = run_main(
        "int a = 0; while (a < 2) { int b = 0; while (b < 2) { int c = 0; while (c < 2) { if (c == 1) break; System.out.println(a + b + c); c++; } b++; } a++; }",
    );
    assert_eq!(out, vec!["0", "1", "1", "2"]);
}

#[test]
fn while_with_postfix_decrement_in_condition() {
    let out = run_main("int n = 4; while (n-- > 1) { System.out.println(n); }");
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn while_with_prefix_increment_in_condition_guard() {
    let out = run_main("int n = 0; while (++n < 4) { System.out.println(n); }");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn while_short_circuit_and_in_condition() {
    let out = run_main(
        "int x = 0; int y = 0; while (x < 3 && (y++ < 5)) { System.out.println(x); x++; }",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn while_short_circuit_or_stops_when_left_true() {
    let out = run_main(
        "int a = 0; while ((a++ == 1) || (a < 0)) { System.out.println(a); if (a > 2) break; }",
    );
    assert_eq!(out, Vec::<String>::new());
}

#[test]
fn do_while_nested_inside_while_menu_stack() {
    let out = run_main(
        "int page = 0; while (page < 2) { int item = 0; do { System.out.println(page + \"-\" + item); item++; } while (item < 2); page++; }",
    );
    assert_eq!(out, vec!["0-0", "0-1", "1-0", "1-1"]);
}

#[test]
fn while_nested_inside_do_while_outer_menu() {
    let out = run_main(
        "int session = 0; do { int step = 0; while (step < 2) { System.out.println(session * 10 + step); step++; } session++; } while (session < 2);",
    );
    assert_eq!(out, vec!["0", "1", "10", "11"]);
}

#[test]
fn while_sentinel_zero_terminates_accumulator() {
    let out = run_main(
        "int[] nums = {2, 3, 0, 9}; int i = 0; int sum = 0; while (nums[i] != 0) { sum += nums[i]; i++; } System.out.println(sum);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn while_complex_range_check_with_or_escape() {
    let out =
        run_main("int v = 8; while (v > 0 && (v < 5 || v == 8)) { System.out.println(v); v--; }");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn nested_break_then_continue_in_while_grid() {
    let out = run_main(
        "int r = 0; while (r < 2) { int c = 0; while (c < 3) { c++; if (c == 2) continue; if (c == 3) break; System.out.println(r * 10 + c); } r++; }",
    );
    assert_eq!(out, vec!["1", "11"]);
}

#[test]
fn do_while_menu_with_default_action_once() {
    let out = run_main(
        "int action = 9; do { System.out.println(\"default\"); action = 0; } while (action != 0);",
    );
    assert_eq!(out, vec!["default"]);
}

#[test]
fn while_doubling_until_threshold() {
    let out = run_main("int n = 1; while (n < 20) { System.out.println(n); n *= 2; }");
    assert_eq!(out, vec!["1", "2", "4", "8", "16"]);
}

#[test]
fn while_with_negated_equality_sentinel() {
    let out = run_main(
        "int[] marks = {70, 80, 90, -99}; int i = 0; while (marks[i] != -99) { System.out.println(marks[i]); i++; }",
    );
    assert_eq!(out, vec!["70", "80", "90"]);
}

#[test]
fn nested_while_break_on_diagonal_match() {
    let out = run_main(
        "int i = 0; while (i < 3) { int j = 0; while (j < 3) { if (i == j && i == 2) break; System.out.println(i * 10 + j); j++; } i++; }",
    );
    assert_eq!(out, vec!["0", "1", "2", "10", "11", "12", "20", "21"]);
}

#[test]
fn do_while_continue_skips_logging_on_second_round() {
    let out = run_main(
        "int n = 0; do { n++; if (n == 2) continue; System.out.println(n); } while (n < 4);",
    );
    assert_eq!(out, vec!["1", "3", "4"]);
}

#[test]
fn while_three_part_condition_with_arithmetic() {
    let out = run_main(
        "int a = 1; int b = 10; while (a < 4 && b > 7 && a + b < 14) { System.out.println(a); a++; b--; }",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn while_reads_chars_until_space_sentinel() {
    let out = run_main(
        r#"String s = "abc def"; int i = 0; while (s.charAt(i) != ' ') { System.out.println(s.charAt(i)); i++; }"#,
    );
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn nested_while_outer_continue_skips_row() {
    let out = run_main(
        "int r = 0; while (r < 3) { r++; if (r == 2) continue; int c = 0; while (c < 2) { System.out.println(r * 10 + c); c++; } }",
    );
    assert_eq!(out, vec!["10", "11", "30", "31"]);
}

#[test]
fn do_while_menu_processes_then_checks_exit() {
    let out = run_main(
        "int running = 1; int count = 0; do { count++; System.out.println(count); running = count == 3 ? 0 : 1; } while (running != 0);",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn while_with_boolean_flag_and_counter() {
    let out = run_main(
        "int i = 0; boolean active = true; while (active && i < 3) { System.out.println(i); i++; if (i == 3) active = false; }",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn while_sentinel_on_decrementing_index() {
    let out = run_main(
        "int[] stack = {0, 0, 5, 0}; int top = 3; while (stack[top] != 0) { System.out.println(stack[top]); top--; }",
    );
    assert_eq!(out, Vec::<String>::new());
}

#[test]
fn nested_while_inner_break_on_sum_threshold() {
    let out = run_main(
        "int r = 0; while (r < 3) { int c = 0; int sum = 0; while (c < 5) { sum += c; if (sum > 3) break; System.out.println(sum); c++; } r++; }",
    );
    assert_eq!(out, vec!["0", "1", "3", "0", "1", "3", "0", "1", "3"]);
}

#[test]
fn do_while_at_least_once_even_when_guard_false() {
    let out = run_main("int x = 10; do { System.out.println(x); } while (x < 5);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn while_not_condition_with_incrementing_counter() {
    let out = run_main("int n = 0; while (!(n >= 3)) { System.out.println(n); n++; }");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn while_nested_break_exits_only_inner_on_column_limit() {
    let out = run_main(
        "int row = 0; while (row < 2) { int col = 0; while (col < 5) { if (col == 2) break; System.out.println(row * 100 + col); col++; } row++; }",
    );
    assert_eq!(out, vec!["0", "1", "100", "101"]);
}

#[test]
fn do_while_menu_nested_break_to_exit_session() {
    let out = run_main(
        "int session = 0; do { session++; int item = 0; while (item < 3) { item++; if (item == 2) break; System.out.println(session + \".\" + item); } if (session == 2) break; } while (session < 5);",
    );
    assert_eq!(out, vec!["1.1", "2.1"]);
}

#[test]
fn while_mixed_int_comparison_with_division() {
    let out = run_main("int n = 16; while (n / 2 >= 2) { System.out.println(n); n /= 2; }");
    assert_eq!(out, vec!["16", "8", "4"]);
}

#[test]
fn while_condition_with_equality_and_less_combined() {
    let out = run_main("int x = 0; while (x < 4 && x != 3) { System.out.println(x); x++; }");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn nested_while_outer_break_on_first_inner_completion() {
    let out = run_main(
        "int a = 0; while (a < 5) { int b = 0; while (b < 2) { System.out.println(a); b++; } break; a++; }",
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn do_while_menu_with_switch_like_if_ladder() {
    let out = run_main(
        "int cmd = 1; do { if (cmd == 1) { System.out.println(\"list\"); } else if (cmd == 2) { System.out.println(\"add\"); } cmd++; } while (cmd < 3);",
    );
    assert_eq!(out, vec!["list", "add"]);
}
