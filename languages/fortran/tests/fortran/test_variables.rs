use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// Fortran: Variable declarations, types, initialization
// ═══════════════════════════════════════════════════════════

#[test]
fn integer_default() {
    let out = run_prints("program t\ninteger :: x = 0\nprint *, x\nend program t\n");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn integer_assign() {
    let out = run_prints("program t\ninteger :: x\nx = 99\nprint *, x\nend program t\n");
    assert_eq!(out, vec!["99"]);
}

#[test]
fn real_assign() {
    let out = run_prints("program t\nreal :: x\nx = 2.5\nprint *, x\nend program t\n");
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn real_init() {
    let out = run_prints("program t\nreal :: pi = 3.14159\nprint *, pi\nend program t\n");
    assert_eq!(out, vec!["3.14159"]);
}

#[test]
fn logical_true() {
    let out = run_prints("program t\nlogical :: flag = .true.\nprint *, flag\nend program t\n");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn logical_false() {
    let out = run_prints("program t\nlogical :: flag = .false.\nprint *, flag\nend program t\n");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_init() {
    // Fortran fixed-length characters pad shorter literals with trailing blanks.
    let out =
        run_prints("program t\ncharacter(len=10) :: s = \"hello\"\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["hello     "]);
}

#[test]
fn character_assign() {
    let out =
        run_prints("program t\ncharacter(len=20) :: s\ns = \"world\"\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["world"]);
}

#[test]
fn multiple_same_type() {
    let out = run_prints(
        "program t\ninteger :: a = 1, b = 2, c = 3\nprint *, a + b + c\nend program t\n",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn parameter_constant() {
    let out = run_prints("program t\ninteger, parameter :: N = 42\nprint *, N\nend program t\n");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn parameter_real() {
    let out =
        run_prints("program t\nreal, parameter :: PI = 3.14159\nprint *, PI\nend program t\n");
    assert_eq!(out, vec!["3.14159"]);
}

#[test]
fn negative_number() {
    let out = run_prints("program t\ninteger :: x = -5\nprint *, x\nend program t\n");
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn reassign_variable() {
    let out = run_prints("program t\ninteger :: x\nx = 10\nx = 20\nprint *, x\nend program t\n");
    assert_eq!(out, vec!["20"]);
}

#[test]
fn double_precision_var() {
    let out = run_prints("program t\ndouble precision :: d = 1.23456789d0\nprint *, nint(d*1_8)\nend program t\n");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn implicit_none_compiles() {
    let out = run_prints("program t\nimplicit none\ninteger :: x = 1\nprint *, x\nend program t\n");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn complex_variable_runtime() {
    let out = run_prints("program t\ncomplex :: c = (1.25, -2.5)\nprint *, nint(real(c)*10)\nprint *, nint(aimag(c)*10)\nend program t\n");
    assert_eq!(out, vec!["12", "-25"]);
}

#[test]
fn character_length_truncation_runtime() {
    let out = run_prints("program t\ncharacter(len=3) :: s = 'hello'\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["hel"]);
}

#[test]
fn integer_kind8_runtime() {
    let out = run_prints("program t\ninteger(kind=8) :: k = 123456789012_8\nprint *, k\nend program t\n");
    assert_eq!(out, vec!["123456789012"]);
}
