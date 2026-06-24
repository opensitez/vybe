use crate::helpers::{run_in_main, run_main};

#[test]
fn switch_int_matches_first_case() {
    let out = run_main(
        "int x = 1; switch (x) { case 1: System.out.println(\"one\"); break; case 2: System.out.println(\"two\"); }",
    );
    assert_eq!(out, vec!["one"]);
}

#[test]
fn switch_int_matches_second_case() {
    let out = run_main(
        "int x = 2; switch (x) { case 1: System.out.println(\"one\"); break; case 2: System.out.println(\"two\"); break; }",
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn switch_int_hits_default_when_unmatched() {
    let out = run_main(
        "int x = 9; switch (x) { case 1: System.out.println(\"one\"); default: System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["other"]);
}

#[test]
fn switch_int_falls_through_without_break() {
    let out = run_main(
        "int x = 1; switch (x) { case 1: System.out.println(\"one\"); case 2: System.out.println(\"two\"); }",
    );
    assert_eq!(out, vec!["one", "two"]);
}

#[test]
fn switch_int_break_stops_fallthrough() {
    let out = run_main(
        "int x = 1; switch (x) { case 1: System.out.println(\"one\"); break; case 2: System.out.println(\"two\"); }",
    );
    assert_eq!(out, vec!["one"]);
}

#[test]
fn switch_string_matches_literal_label() {
    let out = run_main(
        "String op = \"add\"; switch (op) { case \"add\": System.out.println(\"plus\"); break; case \"sub\": System.out.println(\"minus\"); }",
    );
    assert_eq!(out, vec!["plus"]);
}

#[test]
fn switch_string_default_when_label_absent() {
    let out = run_main(
        "String op = \"mul\"; switch (op) { case \"add\": System.out.println(\"plus\"); default: System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["other"]);
}

#[test]
fn switch_enum_constant_selects_matching_arm() {
    let types = r#"
        enum Color { RED, GREEN, BLUE }
    "#;
    let out = run_in_main(
        "Color c = Color.BLUE; switch (c) { case RED: System.out.println(\"r\"); break; case GREEN: System.out.println(\"g\"); break; case BLUE: System.out.println(\"b\"); break; }",
        types,
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn switch_enum_default_arm_runs_for_unhandled_mapping() {
    let types = r#"
        enum Tier { LOW, HIGH }
    "#;
    let out = run_in_main(
        "Tier t = Tier.HIGH; switch (t) { case LOW: System.out.println(\"l\"); break; default: System.out.println(\"h\"); }",
        types,
    );
    assert_eq!(out, vec!["h"]);
}

#[test]
fn switch_expression_arrow_yields_matching_string() {
    let out = run_main(
        "int code = 2; String label = switch (code) { case 1 -> \"low\"; case 2 -> \"mid\"; default -> \"high\"; }; System.out.println(label);",
    );
    assert_eq!(out, vec!["mid"]);
}

#[test]
fn switch_expression_arrow_uses_default_arm() {
    let out = run_main(
        "int code = 99; String label = switch (code) { case 1 -> \"low\"; case 2 -> \"mid\"; default -> \"high\"; }; System.out.println(label);",
    );
    assert_eq!(out, vec!["high"]);
}

#[test]
fn switch_arrow_statement_runs_single_action() {
    let out = run_main(
        "int n = 3; switch (n) { case 3 -> System.out.println(\"three\"); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["three"]);
}

#[test]
fn switch_int_case_zero_is_reachable() {
    let out = run_main(
        "int x = 0; switch (x) { case 0: System.out.println(\"zero\"); break; default: System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["zero"]);
}

#[test]
fn switch_int_negative_value_matches_case() {
    let out = run_main(
        "int x = -1; switch (x) { case -1: System.out.println(\"neg\"); break; default: System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["neg"]);
}

#[test]
fn switch_case_body_runs_multiple_statements() {
    let out = run_main(
        "int x = 2; switch (x) { case 2: System.out.println(\"a\"); System.out.println(\"b\"); break; default: System.out.println(\"z\"); }",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn switch_fallthrough_runs_three_consecutive_cases() {
    let out = run_main(
        "int x = 1; switch (x) { case 1: System.out.println(\"1\"); case 2: System.out.println(\"2\"); case 3: System.out.println(\"3\"); }",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn switch_default_only_arm_executes() {
    let out = run_main(
        "int x = 50; switch (x) { default: System.out.println(\"only\"); }",
    );
    assert_eq!(out, vec!["only"]);
}

#[test]
fn switch_string_empty_literal_matches_case() {
    let out = run_main(
        "String s = \"\"; switch (s) { case \"\": System.out.println(\"empty\"); break; default: System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["empty"]);
}

#[test]
fn switch_int_break_before_default_skips_default() {
    let out = run_main(
        "int x = 1; switch (x) { case 1: System.out.println(\"hit\"); break; default: System.out.println(\"miss\"); }",
    );
    assert_eq!(out, vec!["hit"]);
}

#[test]
fn switch_expression_arrow_returns_integer_score() {
    let out = run_main(
        "int rank = 3; int score = switch (rank) { case 1 -> 100; case 2 -> 80; default -> 50; }; System.out.println(score);",
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn switch_nested_inside_if_uses_outer_condition() {
    let out = run_main(
        "int x = 2; if (x > 0) { switch (x) { case 2: System.out.println(\"two\"); break; default: System.out.println(\"other\"); } }",
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn switch_enum_arrow_rule_prints_label() {
    let types = r#"
        enum Mode { ON, OFF }
    "#;
    let out = run_in_main(
        "Mode m = Mode.ON; switch (m) { case ON -> System.out.println(\"on\"); case OFF -> System.out.println(\"off\"); }",
        types,
    );
    assert_eq!(out, vec!["on"]);
}

#[test]
fn switch_int_multiple_labels_share_one_arm() {
    let out = run_main(
        "int x = 2; switch (x) { case 1, 2: System.out.println(\"small\"); break; default: System.out.println(\"big\"); }",
    );
    assert_eq!(out, vec!["small"]);
}

#[test]
fn switch_string_second_literal_in_multi_label_arm() {
    let out = run_main(
        "String key = \"b\"; switch (key) { case \"a\", \"b\": System.out.println(\"ab\"); break; default: System.out.println(\"z\"); }",
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn switch_int_no_matching_case_and_no_default_is_silent() {
    let out = run_main(
        "int x = 7; switch (x) { case 1: System.out.println(\"one\"); case 2: System.out.println(\"two\"); } System.out.println(\"done\");",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn switch_expression_first_arm_selected_for_code_one() {
    let out = run_main(
        "int code = 1; String out = switch (code) { case 1 -> \"one\"; case 2 -> \"two\"; default -> \"many\"; }; System.out.println(out);",
    );
    assert_eq!(out, vec!["one"]);
}

#[test]
fn switch_int_fall_from_case_two_into_case_three() {
    let out = run_main(
        "int x = 2; switch (x) { case 1: System.out.println(\"skip\"); case 2: System.out.println(\"two\"); case 3: System.out.println(\"three\"); break; }",
    );
    assert_eq!(out, vec!["two", "three"]);
}

#[test]
fn switch_string_non_matching_case_falls_to_default() {
    let out = run_main(
        "String s = \"z\"; switch (s) { case \"a\": System.out.println(\"a\"); break; case \"b\": System.out.println(\"b\"); break; default: System.out.println(\"d\"); }",
    );
    assert_eq!(out, vec!["d"]);
}

#[test]
fn switch_arrow_block_with_multiple_statements() {
    let out = run_main(
        "int n = 4; switch (n) { case 4 -> { System.out.println(\"a\"); System.out.println(\"b\"); } default -> System.out.println(\"x\"); }",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn switch_int_variable_reassigned_before_switch() {
    let out = run_main(
        "int x = 1; x = 3; switch (x) { case 3: System.out.println(\"three\"); break; default: System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["three"]);
}
