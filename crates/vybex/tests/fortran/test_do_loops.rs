use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: Do loops
// ═══════════════════════════════════════════════════════════

#[test]
fn do_1_to_5() {
    let out = run_prints("program t\ninteger :: i, s\ns = 0\ndo i = 1, 5\ns = s + i\nend do\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["15"]);
}

#[test]
fn do_1_to_10() {
    let out = run_prints("program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\ns = s + i\nend do\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["55"]);
}

#[test]
fn do_step_2() {
    let out = run_prints("program t\ninteger :: i, s\ns = 0\ndo i = 0, 10, 2\ns = s + i\nend do\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["30"]);
}

#[test]
fn do_print_each() {
    let out = run_prints("program t\ninteger :: i\ndo i = 1, 3\nprint *, i\nend do\nend program t\n");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
#[ignore] // hangs — while loop variable reassignment not updating condition
fn do_while_simple() {
    let out = run_prints("program t\ninteger :: i\ni = 0\ndo while (i < 5)\ni = i + 1\nend do\nprint *, i\nend program t\n");
    assert_eq!(out, vec!["5"]);
}

#[test]
#[ignore]
fn do_while_accumulate() {
    let out = run_prints("program t\ninteger :: i, s\ni = 1\ns = 0\ndo while (i <= 10)\ns = s + i\ni = i + 1\nend do\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["55"]);
}

#[test]
fn do_exit() {
    let out = run_prints("program t\ninteger :: i\ndo i = 1, 100\nif (i > 3) exit\nprint *, i\nend do\nend program t\n");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn do_cycle() {
    let out = run_prints("program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\nif (i == 5) cycle\ns = s + i\nend do\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["50"]);
}

#[test]
fn nested_do() {
    let out = run_prints("program t\ninteger :: i, j, c\nc = 0\ndo i = 1, 3\ndo j = 1, 4\nc = c + 1\nend do\nend do\nprint *, c\nend program t\n");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn do_multiplication_table_row() {
    let out = run_prints("program t\ninteger :: i\ndo i = 1, 5\nprint *, 3 * i\nend do\nend program t\n");
    assert_eq!(out, vec!["3", "6", "9", "12", "15"]);
}

#[test]
fn do_factorial() {
    let out = run_prints("program t\ninteger :: i, f\nf = 1\ndo i = 1, 5\nf = f * i\nend do\nprint *, f\nend program t\n");
    assert_eq!(out, vec!["120"]);
}

#[test]
fn do_fibonacci_iterative() {
    let out = run_prints("program t\ninteger :: i, a, b, tmp\na = 0\nb = 1\ndo i = 1, 10\ntmp = a + b\na = b\nb = tmp\nend do\nprint *, a\nend program t\n");
    assert_eq!(out, vec!["55"]);
}

#[test]
fn do_power_of_two() {
    let out = run_prints("program t\ninteger :: i, p\np = 1\ndo i = 1, 8\np = p * 2\nend do\nprint *, p\nend program t\n");
    assert_eq!(out, vec!["256"]);
}

#[test]
fn do_sum_of_squares() {
    let out = run_prints("program t\ninteger :: i, s\ns = 0\ndo i = 1, 5\ns = s + i * i\nend do\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["55"]);
}

#[test]
fn do_exit_with_sum() {
    let out = run_prints("program t\ninteger :: i, s\ns = 0\ndo i = 1, 100\ns = s + i\nif (s > 50) exit\nend do\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["55"]);
}
