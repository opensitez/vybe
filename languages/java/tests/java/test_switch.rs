use crate::helpers::run_main;

#[test]
fn switch_falls_through_without_break() {
    let out = run_main(
        "int x = 1; switch (x) { case 1: System.out.println(\"one\"); case 2: System.out.println(\"two\"); }",
    );
    assert_eq!(out, vec!["one", "two"]);
}

#[test]
fn switch_break_prevents_fallthrough() {
    let out = run_main(
        "int x = 1; switch (x) { case 1: System.out.println(\"one\"); break; case 2: System.out.println(\"two\"); }",
    );
    assert_eq!(out, vec!["one"]);
}

#[test]
fn switch_default_when_no_case_matches() {
    let out = run_main(
        "int x = 99; switch (x) { case 1: System.out.println(\"one\"); default: System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["other"]);
}

#[test]
fn switch_on_string_matches_label() {
    let out = run_main(
        "String op = \"add\"; switch (op) { case \"add\": System.out.println(\"plus\"); break; case \"sub\": System.out.println(\"minus\"); }",
    );
    assert_eq!(out, vec!["plus"]);
}

#[test]
fn switch_expression_yields_value() {
    let out = run_main(
        "int code = 2; String label = switch (code) { case 1 -> \"low\"; case 2 -> \"mid\"; default -> \"high\"; }; System.out.println(label);",
    );
    assert_eq!(out, vec!["mid"]);
}
