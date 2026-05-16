use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: Print and I/O
// ═══════════════════════════════════════════════════════════

#[test]
fn print_string() {
    let out = run_prints("program t\nprint *, \"Hello\"\nend program t\n");
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn print_integer() {
    let out = run_prints("program t\nprint *, 42\nend program t\n");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn print_real() {
    let out = run_prints("program t\nprint *, 3.14\nend program t\n");
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn print_logical_true() {
    let out = run_prints("program t\nprint *, .true.\nend program t\n");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn print_logical_false() {
    let out = run_prints("program t\nprint *, .false.\nend program t\n");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn print_expression() {
    let out = run_prints("program t\nprint *, 2 + 3\nend program t\n");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn print_variable() {
    let out = run_prints("program t\ninteger :: x = 99\nprint *, x\nend program t\n");
    assert_eq!(out, vec!["99"]);
}

#[test]
fn print_multiple_values() {
    let out = run_prints("program t\nprint *, \"x =\", 42\nend program t\n");
    assert_eq!(out, vec!["x = 42"]);
}

#[test]
fn print_multiple_lines() {
    let out = run_prints("program t\nprint *, \"a\"\nprint *, \"b\"\nprint *, \"c\"\nend program t\n");
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn print_negative() {
    let out = run_prints("program t\nprint *, -10\nend program t\n");
    assert_eq!(out, vec!["-10"]);
}
