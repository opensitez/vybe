use crate::helpers::run_main;

#[test]
fn if_true_branch_executes_when_condition_holds() {
    let out = run_main("int x = 10; if (x > 5) { System.out.println(\"yes\"); }");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_false_branch_skipped_when_condition_fails() {
    let out = run_main(
        "int x = 1; if (x > 5) { System.out.println(\"yes\"); } System.out.println(\"done\");",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn if_else_selects_if_when_true() {
    let out = run_main(
        "int x = 8; if (x > 5) { System.out.println(\"big\"); } else { System.out.println(\"small\"); }",
    );
    assert_eq!(out, vec!["big"]);
}

#[test]
fn if_else_selects_else_when_false() {
    let out = run_main(
        "int x = 2; if (x > 5) { System.out.println(\"big\"); } else { System.out.println(\"small\"); }",
    );
    assert_eq!(out, vec!["small"]);
}

#[test]
fn else_if_chain_picks_first_matching_branch() {
    let out = run_main(
        "int score = 75; if (score >= 90) { System.out.println(\"A\"); } else if (score >= 70) { System.out.println(\"B\"); } else { System.out.println(\"C\"); }",
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn else_if_chain_falls_through_to_else() {
    let out = run_main(
        "int score = 55; if (score >= 90) { System.out.println(\"A\"); } else if (score >= 70) { System.out.println(\"B\"); } else { System.out.println(\"C\"); }",
    );
    assert_eq!(out, vec!["C"]);
}

#[test]
fn else_if_chain_stops_at_first_true() {
    let out = run_main(
        "int n = 95; if (n >= 90) { System.out.println(\"A\"); } else if (n >= 80) { System.out.println(\"B\"); } else if (n >= 70) { System.out.println(\"C\"); } else { System.out.println(\"F\"); }",
    );
    assert_eq!(out, vec!["A"]);
}

#[test]
fn else_if_chain_middle_branch_wins() {
    let out = run_main(
        "int n = 85; if (n >= 90) { System.out.println(\"A\"); } else if (n >= 80) { System.out.println(\"B\"); } else if (n >= 70) { System.out.println(\"C\"); } else { System.out.println(\"F\"); }",
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn nested_if_both_conditions_true() {
    let out = run_main(
        "int a = 10; int b = 3; if (a > 5) { if (b < 5) { System.out.println(\"inner\"); } }",
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn nested_if_outer_false_skips_inner() {
    let out = run_main(
        "int a = 1; int b = 3; if (a > 5) { if (b < 5) { System.out.println(\"inner\"); } } System.out.println(\"end\");",
    );
    assert_eq!(out, vec!["end"]);
}

#[test]
fn nested_if_inner_false_skips_inner_body() {
    let out = run_main(
        "int a = 10; int b = 9; if (a > 5) { if (b < 5) { System.out.println(\"inner\"); } else { System.out.println(\"else-inner\"); } }",
    );
    assert_eq!(out, vec!["else-inner"]);
}

#[test]
fn nested_if_else_if_chain() {
    let out = run_main(
        "int x = 15; if (x < 10) { System.out.println(\"low\"); } else { if (x < 20) { System.out.println(\"mid\"); } else { System.out.println(\"high\"); } }",
    );
    assert_eq!(out, vec!["mid"]);
}

#[test]
fn boolean_literal_true_enters_if() {
    let out = run_main("if (true) { System.out.println(\"on\"); }");
    assert_eq!(out, vec!["on"]);
}

#[test]
fn boolean_literal_false_skips_if() {
    let out = run_main("if (false) { System.out.println(\"on\"); } System.out.println(\"off\");");
    assert_eq!(out, vec!["off"]);
}

#[test]
fn boolean_variable_controls_branch() {
    let out = run_main(
        "boolean ready = true; if (ready) { System.out.println(\"go\"); } else { System.out.println(\"wait\"); }",
    );
    assert_eq!(out, vec!["go"]);
}

#[test]
fn logical_not_inverts_condition() {
    let out = run_main(
        "boolean flag = false; if (!flag) { System.out.println(\"inverted\"); }",
    );
    assert_eq!(out, vec!["inverted"]);
}

#[test]
fn int_equality_true_branch() {
    let out = run_main(
        "int a = 7; int b = 7; if (a == b) { System.out.println(\"equal\"); } else { System.out.println(\"diff\"); }",
    );
    assert_eq!(out, vec!["equal"]);
}

#[test]
fn int_inequality_detects_difference() {
    let out = run_main(
        "int a = 7; int b = 8; if (a != b) { System.out.println(\"diff\"); } else { System.out.println(\"equal\"); }",
    );
    assert_eq!(out, vec!["diff"]);
}

#[test]
fn int_less_than_comparison() {
    let out = run_main("int x = 3; if (x < 5) { System.out.println(\"below\"); }");
    assert_eq!(out, vec!["below"]);
}

#[test]
fn int_greater_than_comparison() {
    let out = run_main("int x = 9; if (x > 5) { System.out.println(\"above\"); }");
    assert_eq!(out, vec!["above"]);
}

#[test]
fn int_less_or_equal_boundary() {
    let out = run_main(
        "int x = 5; if (x <= 5) { System.out.println(\"at-most\"); } else { System.out.println(\"over\"); }",
    );
    assert_eq!(out, vec!["at-most"]);
}

#[test]
fn int_greater_or_equal_boundary() {
    let out = run_main(
        "int x = 5; if (x >= 5) { System.out.println(\"at-least\"); } else { System.out.println(\"under\"); }",
    );
    assert_eq!(out, vec!["at-least"]);
}

#[test]
fn comparing_zero_in_condition() {
    let out = run_main(
        "int x = 0; if (x == 0) { System.out.println(\"zero\"); } else { System.out.println(\"nonzero\"); }",
    );
    assert_eq!(out, vec!["zero"]);
}

#[test]
fn comparing_negative_integers() {
    let out = run_main(
        "int x = -3; if (x < 0) { System.out.println(\"negative\"); } else { System.out.println(\"nonneg\"); }",
    );
    assert_eq!(out, vec!["negative"]);
}

#[test]
fn string_equals_true_branch() {
    let out = run_main(
        r#"String a = "java"; String b = "java"; if (a.equals(b)) { System.out.println("match"); } else { System.out.println("mismatch"); }"#,
    );
    assert_eq!(out, vec!["match"]);
}

#[test]
fn string_equals_false_branch() {
    let out = run_main(
        r#"String a = "java"; String b = "kotlin"; if (a.equals(b)) { System.out.println("match"); } else { System.out.println("mismatch"); }"#,
    );
    assert_eq!(out, vec!["mismatch"]);
}

#[test]
fn string_not_equal_via_negated_equals() {
    let out = run_main(
        r#"String a = "foo"; String b = "bar"; if (!a.equals(b)) { System.out.println("different"); }"#,
    );
    assert_eq!(out, vec!["different"]);
}

#[test]
fn string_is_empty_condition() {
    let out = run_main(
        r#"String s = ""; if (s.isEmpty()) { System.out.println("empty"); } else { System.out.println("filled"); }"#,
    );
    assert_eq!(out, vec!["empty"]);
}

#[test]
fn string_length_in_condition() {
    let out = run_main(
        r#"String s = "hello"; if (s.length() == 5) { System.out.println("five"); }"#,
    );
    assert_eq!(out, vec!["five"]);
}

#[test]
fn logical_and_requires_both_operands_true() {
    let out = run_main(
        "boolean a = true; boolean b = true; if (a && b) { System.out.println(\"both\"); } else { System.out.println(\"not-both\"); }",
    );
    assert_eq!(out, vec!["both"]);
}

#[test]
fn logical_and_fails_when_one_operand_false() {
    let out = run_main(
        "boolean a = true; boolean b = false; if (a && b) { System.out.println(\"both\"); } else { System.out.println(\"not-both\"); }",
    );
    assert_eq!(out, vec!["not-both"]);
}

#[test]
fn logical_or_succeeds_with_one_true_operand() {
    let out = run_main(
        "boolean a = false; boolean b = true; if (a || b) { System.out.println(\"either\"); } else { System.out.println(\"neither\"); }",
    );
    assert_eq!(out, vec!["either"]);
}

#[test]
fn logical_or_fails_when_both_operands_false() {
    let out = run_main(
        "boolean a = false; boolean b = false; if (a || b) { System.out.println(\"either\"); } else { System.out.println(\"neither\"); }",
    );
    assert_eq!(out, vec!["neither"]);
}

#[test]
fn compound_int_and_boolean_condition() {
    let out = run_main(
        "int age = 20; boolean active = true; if (age >= 18 && active) { System.out.println(\"adult-active\"); }",
    );
    assert_eq!(out, vec!["adult-active"]);
}

#[test]
fn compound_or_with_int_range_check() {
    let out = run_main(
        "int x = 2; if (x < 0 || x > 10) { System.out.println(\"out\"); } else { System.out.println(\"in\"); }",
    );
    assert_eq!(out, vec!["in"]);
}

#[test]
fn if_without_braces_single_statement() {
    let out = run_main("int x = 1; if (x == 1) System.out.println(\"one\");");
    assert_eq!(out, vec!["one"]);
}

#[test]
fn else_if_four_way_grade_scale() {
    let out = run_main(
        "int g = 72; if (g >= 90) { System.out.println(\"A\"); } else if (g >= 80) { System.out.println(\"B\"); } else if (g >= 70) { System.out.println(\"C\"); } else if (g >= 60) { System.out.println(\"D\"); } else { System.out.println(\"F\"); }",
    );
    assert_eq!(out, vec!["C"]);
}
