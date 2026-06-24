use crate::helpers::run_main;

#[test]
fn for_loop_counts_up_to_exclusive_bound() {
    let out = run_main(
        "for (int i = 0; i < 3; i++) { System.out.println(i); }",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn while_loop_runs_until_condition_false() {
    let out = run_main(
        "int n = 3; while (n > 0) { System.out.println(n); n--; }",
    );
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn do_while_executes_body_at_least_once() {
    let out = run_main(
        "int n = 0; do { System.out.println(n); n++; } while (n < 1);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn enhanced_for_iterates_array_elements() {
    let out = run_main(
        "int[] nums = {10, 20, 30}; for (int v : nums) { System.out.println(v); }",
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn break_exits_loop_early() {
    let out = run_main(
        "for (int i = 0; i < 10; i++) { if (i == 3) break; System.out.println(i); }",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn continue_skips_current_iteration() {
    let out = run_main(
        "for (int i = 0; i < 5; i++) { if (i % 2 == 0) continue; System.out.println(i); }",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn nested_loops_produce_cartesian_pairs() {
    let out = run_main(
        "for (int r = 0; r < 2; r++) { for (int c = 0; c < 2; c++) { System.out.println(r * 10 + c); } }",
    );
    assert_eq!(out, vec!["0", "1", "10", "11"]);
}
