use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: Do loops
// ═══════════════════════════════════════════════════════════

#[test]
fn do_1_to_5() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 5\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn do_1_to_10() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn do_step_2() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 0, 10, 2\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn do_print_each() {
    let out =
        run_prints("program t\ninteger :: i\ndo i = 1, 3\nprint *, i\nend do\nend program t\n");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn do_while_simple() {
    let out = run_prints(
        "program t\ninteger :: i\ni = 0\ndo while (i < 5)\ni = i + 1\nend do\nprint *, i\nend program t\n",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn do_while_accumulate() {
    let out = run_prints(
        "program t\ninteger :: i, s\ni = 1\ns = 0\ndo while (i <= 10)\ns = s + i\ni = i + 1\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn do_exit() {
    let out = run_prints(
        "program t\ninteger :: i\ndo i = 1, 100\nif (i > 3) exit\nprint *, i\nend do\nend program t\n",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn do_cycle() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\nif (i == 5) cycle\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn nested_do() {
    let out = run_prints(
        "program t\ninteger :: i, j, c\nc = 0\ndo i = 1, 3\ndo j = 1, 4\nc = c + 1\nend do\nend do\nprint *, c\nend program t\n",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn do_multiplication_table_row() {
    let out =
        run_prints("program t\ninteger :: i\ndo i = 1, 5\nprint *, 3 * i\nend do\nend program t\n");
    assert_eq!(out, vec!["3", "6", "9", "12", "15"]);
}

#[test]
fn do_factorial() {
    let out = run_prints(
        "program t\ninteger :: i, f\nf = 1\ndo i = 1, 5\nf = f * i\nend do\nprint *, f\nend program t\n",
    );
    assert_eq!(out, vec!["120"]);
}

#[test]
fn do_fibonacci_iterative() {
    let out = run_prints(
        "program t\ninteger :: i, a, b, tmp\na = 0\nb = 1\ndo i = 1, 10\ntmp = a + b\na = b\nb = tmp\nend do\nprint *, a\nend program t\n",
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn do_power_of_two() {
    let out = run_prints(
        "program t\ninteger :: i, p\np = 1\ndo i = 1, 8\np = p * 2\nend do\nprint *, p\nend program t\n",
    );
    assert_eq!(out, vec!["256"]);
}

#[test]
fn do_sum_of_squares() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 5\ns = s + i * i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn do_exit_with_sum() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 100\ns = s + i\nif (s > 50) exit\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn bare_do_loop_exits() {
    let out = run_prints(
        "program t\ninteger :: i\ni = 0\ndo\ni = i + 1\nif (i >= 3) exit\nend do\nprint *, i\nend program t\n",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn do_descending_step() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 6, 2, -2\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn do_empty_range_skips_body() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 5, 1\ns = s + 1\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn do_named_loop_exit() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\nouter: do i = 1, 5\n    if (i == 3) exit outer\n    s = s + 1\nend do outer\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn do_start_equals_end_runs_once() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 4, 4\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn do_start_equals_end_negative_step_runs_once() {
    let out =
        run_prints("program t\ninteger :: i, s\ns = 0\ndo i = -3, -3, -2\ns = s + i\nend do\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["-3"]);
}

#[test]
fn do_bound_mutation_is_ignored() {
    let out = run_prints(
        "program t\ninteger :: i, s, bound\nbound = 4\ns = 0\ndo i = 1, bound\nif (i == 2) bound = 10\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn do_step_mutation_is_ignored() {
    let out = run_prints(
        "program t\ninteger :: i, s, step\nstep = 2\ns = 0\ndo i = 1, 10, step\nif (i == 3) step = 1\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn do_labeled_continue_syntax() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo 55 i = 1, 4\ns = s + i\n55 continue\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn do_loop_var_value_after_completion() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 4\ns = s + i\nend do\nprint *, i\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["5", "10"]);
}

#[test]
fn do_named_cycle_skips_named_iteration() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\nouter: do i = 1, 5\n    if (i == 3) cycle outer\n    s = s + i\nend do outer\nprint *, s\nprint *, i\nend program t\n",
    );
    assert_eq!(out, vec!["12", "6"]);
}

#[test]
fn do_with_expression_bounds_and_step() {
    let out = run_prints(
        "program t\ninteger :: i, s\ninteger :: first, last, jump\nfirst = 1\nlast = 10\njump = 3\ns = 0\ndo i = first + 1, last - 1, 1 + jump\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn do_start_is_mutated_inside_loop_but_ignored() {
    let out = run_prints(
        "program t\ninteger :: i, s, start, finish\nstart = 1\nfinish = 3\ns = 0\ndo i = start, finish\n    if (i == 1) start = 99\n    s = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn do_negative_step_with_descending_bounds() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 5, 2, -1\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn do_positive_step_with_negative_direction_is_ignored() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 4, -1\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["0"]);
}
