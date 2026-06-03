use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: Select case
// ═══════════════════════════════════════════════════════════

#[test]
fn select_case_1() {
    let out = run_prints(
        "program t\ninteger :: x = 1\nselect case (x)\ncase (1)\nprint *, \"one\"\ncase (2)\nprint *, \"two\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["one"]);
}

#[test]
fn select_case_2() {
    let out = run_prints(
        "program t\ninteger :: x = 2\nselect case (x)\ncase (1)\nprint *, \"one\"\ncase (2)\nprint *, \"two\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn select_case_default() {
    let out = run_prints(
        "program t\ninteger :: x = 99\nselect case (x)\ncase (1)\nprint *, \"one\"\ncase default\nprint *, \"default\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["default"]);
}

#[test]
fn select_case_no_match_no_default() {
    let out = run_prints(
        "program t\ninteger :: x = 99\nselect case (x)\ncase (1)\nprint *, \"one\"\nend select\nprint *, \"done\"\nend program t\n",
    );
    assert_eq!(out, vec!["done"]);
}
