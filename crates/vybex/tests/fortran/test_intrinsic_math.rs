use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// Fortran: Intrinsic math functions
// ═══════════════════════════════════════════════════════════

#[test]
fn abs_positive() {
    let out = run_prints("program t\nprint *, abs(42)\nend program t\n");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn abs_negative() {
    let out = run_prints("program t\nprint *, abs(-42)\nend program t\n");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn sqrt_25() {
    let out = run_prints("program t\nprint *, sqrt(25.0)\nend program t\n");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn sqrt_100() {
    let out = run_prints("program t\nprint *, sqrt(100.0)\nend program t\n");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn min_two() {
    let out = run_prints("program t\nprint *, min(3, 7)\nend program t\n");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn max_two() {
    let out = run_prints("program t\nprint *, max(3, 7)\nend program t\n");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn sin_zero() {
    compile_ok("program t\nreal :: x\nx = sin(0.0)\nprint *, x\nend program t\n");
}

#[test]
fn cos_zero() {
    compile_ok("program t\nreal :: x\nx = cos(0.0)\nprint *, x\nend program t\n");
}

#[test]
fn exp_one() {
    compile_ok("program t\nreal :: x\nx = exp(1.0)\nprint *, x\nend program t\n");
}

#[test]
fn log_e() {
    compile_ok("program t\nreal :: x\nx = log(2.718)\nprint *, x\nend program t\n");
}

#[test]
fn floor_positive() {
    let out = run_prints("program t\nprint *, floor(3.7)\nend program t\n");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn ceiling_positive() {
    let out = run_prints("program t\nprint *, ceiling(3.2)\nend program t\n");
    assert_eq!(out, vec!["4"]);
}
