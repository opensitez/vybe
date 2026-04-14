use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: Arithmetic expressions
// ═══════════════════════════════════════════════════════════

#[test]
fn add_integers() {
    let out = run_prints("program t\nprint *, 3 + 4\nend program t\n");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn subtract_integers() {
    let out = run_prints("program t\nprint *, 10 - 3\nend program t\n");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn multiply_integers() {
    let out = run_prints("program t\nprint *, 6 * 7\nend program t\n");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn divide_reals() {
    let out = run_prints("program t\nprint *, 10.0 / 4.0\nend program t\n");
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn power_operator() {
    let out = run_prints("program t\nprint *, 2 ** 10\nend program t\n");
    assert_eq!(out, vec!["1024"]);
}

#[test]
fn power_small() {
    let out = run_prints("program t\nprint *, 3 ** 3\nend program t\n");
    assert_eq!(out, vec!["27"]);
}

#[test]
fn add_reals() {
    let out = run_prints("program t\nprint *, 1.5 + 2.5\nend program t\n");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn compound_expression() {
    let out = run_prints("program t\nprint *, 2 + 3 * 4\nend program t\n");
    assert_eq!(out, vec!["14"]);
}

#[test]
fn parenthesized() {
    let out = run_prints("program t\nprint *, (2 + 3) * 4\nend program t\n");
    assert_eq!(out, vec!["20"]);
}

#[test]
fn unary_minus() {
    let out = run_prints("program t\ninteger :: x = 5\nprint *, -x\nend program t\n");
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn add_variables() {
    let out = run_prints("program t\ninteger :: a = 10, b = 20\nprint *, a + b\nend program t\n");
    assert_eq!(out, vec!["30"]);
}

#[test]
fn subtract_variables() {
    let out = run_prints("program t\ninteger :: a = 10, b = 3\nprint *, a - b\nend program t\n");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn multiply_variables() {
    let out = run_prints("program t\ninteger :: a = 6, b = 7\nprint *, a * b\nend program t\n");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn complex_arithmetic() {
    let out = run_prints("program t\nprint *, 2 ** 3 + 1\nend program t\n");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn assign_expression() {
    let out = run_prints("program t\ninteger :: x\nx = 2 + 3 * 4\nprint *, x\nend program t\n");
    assert_eq!(out, vec!["14"]);
}
