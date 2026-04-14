use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: If/then/else blocks
// ═══════════════════════════════════════════════════════════

#[test]
fn if_true() {
    let out = run_prints("program t\nif (1 > 0) then\nprint *, \"yes\"\nend if\nend program t\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_false() {
    let out = run_prints("program t\nif (0 > 1) then\nprint *, \"no\"\nend if\nprint *, \"done\"\nend program t\n");
    assert_eq!(out, vec!["done"]);
}

#[test]
fn if_else_true() {
    let out = run_prints("program t\nif (5 > 3) then\nprint *, \"big\"\nelse\nprint *, \"small\"\nend if\nend program t\n");
    assert_eq!(out, vec!["big"]);
}

#[test]
fn if_else_false() {
    let out = run_prints("program t\nif (1 > 5) then\nprint *, \"big\"\nelse\nprint *, \"small\"\nend if\nend program t\n");
    assert_eq!(out, vec!["small"]);
}

#[test]
fn if_elseif() {
    let out = run_prints("program t\ninteger :: x = 2\nif (x == 1) then\nprint *, \"one\"\nelse if (x == 2) then\nprint *, \"two\"\nelse\nprint *, \"other\"\nend if\nend program t\n");
    assert_eq!(out, vec!["two"]);
}

#[test]
fn if_elseif_default() {
    let out = run_prints("program t\ninteger :: x = 99\nif (x == 1) then\nprint *, \"one\"\nelse if (x == 2) then\nprint *, \"two\"\nelse\nprint *, \"other\"\nend if\nend program t\n");
    assert_eq!(out, vec!["other"]);
}

#[test]
fn if_gt() {
    let out = run_prints("program t\nif (5 > 3) then\nprint *, \"yes\"\nend if\nend program t\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_lt() {
    let out = run_prints("program t\nif (3 < 5) then\nprint *, \"yes\"\nend if\nend program t\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_eq() {
    let out = run_prints("program t\nif (3 == 3) then\nprint *, \"equal\"\nend if\nend program t\n");
    assert_eq!(out, vec!["equal"]);
}

#[test]
fn if_ne() {
    let out = run_prints("program t\nif (3 /= 4) then\nprint *, \"not equal\"\nend if\nend program t\n");
    assert_eq!(out, vec!["not equal"]);
}

#[test]
fn if_ge() {
    let out = run_prints("program t\nif (5 >= 5) then\nprint *, \"yes\"\nend if\nend program t\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_le() {
    let out = run_prints("program t\nif (3 <= 3) then\nprint *, \"yes\"\nend if\nend program t\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_with_variable() {
    let out = run_prints("program t\ninteger :: score = 85\nif (score >= 90) then\nprint *, \"A\"\nelse if (score >= 80) then\nprint *, \"B\"\nelse\nprint *, \"C\"\nend if\nend program t\n");
    assert_eq!(out, vec!["B"]);
}

#[test]
fn if_logical_and() {
    let out = run_prints("program t\nif (1 > 0 .and. 2 > 1) then\nprint *, \"both\"\nend if\nend program t\n");
    assert_eq!(out, vec!["both"]);
}

#[test]
fn if_logical_or() {
    let out = run_prints("program t\nif (1 > 5 .or. 2 > 1) then\nprint *, \"either\"\nend if\nend program t\n");
    assert_eq!(out, vec!["either"]);
}

#[test]
fn if_logical_not() {
    let out = run_prints("program t\nif (.not. (1 > 5)) then\nprint *, \"negated\"\nend if\nend program t\n");
    assert_eq!(out, vec!["negated"]);
}

#[test]
fn nested_if() {
    let out = run_prints("program t\nif (1 > 0) then\nif (2 > 1) then\nprint *, \"nested\"\nend if\nend if\nend program t\n");
    assert_eq!(out, vec!["nested"]);
}

#[test]
fn single_line_if_print() {
    let out = run_prints("program t\nif (1 > 0) print *, \"inline\"\nend program t\n");
    assert_eq!(out, vec!["inline"]);
}

#[test]
fn if_after_assignment() {
    let out = run_prints("program t\ninteger :: x = 10\nx = x + 5\nif (x == 15) then\nprint *, \"correct\"\nend if\nend program t\n");
    assert_eq!(out, vec!["correct"]);
}

#[test]
fn if_multiple_statements_in_body() {
    let out = run_prints("program t\nif (1 > 0) then\nprint *, \"a\"\nprint *, \"b\"\nprint *, \"c\"\nend if\nend program t\n");
    assert_eq!(out, vec!["a", "b", "c"]);
}
