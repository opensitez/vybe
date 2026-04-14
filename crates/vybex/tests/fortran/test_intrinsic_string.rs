use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: Intrinsic string functions
// ═══════════════════════════════════════════════════════════

#[test]
fn len_string() {
    let out = run_prints("program t\ncharacter(len=10) :: s = \"hello\"\nprint *, len(s)\nend program t\n");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn trim_string() {
    let out = run_prints("program t\nprint *, trim(\"  hello  \")\nend program t\n");
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn len_trim_string() {
    let out = run_prints("program t\ncharacter(len=20) :: s = \"hello\"\nprint *, len_trim(s)\nend program t\n");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn repeat_string() {
    let out = run_prints("program t\nprint *, repeat(\"ab\", 3)\nend program t\n");
    assert_eq!(out, vec!["ababab"]);
}

#[test]
fn adjustl_string() {
    let out = run_prints("program t\nprint *, adjustl(\"  hello\")\nend program t\n");
    assert_eq!(out, vec!["hello"]);
}
