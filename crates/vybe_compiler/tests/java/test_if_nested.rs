use crate::helpers::run_main;

#[test]
fn dangling_else_binds_to_nearest_inner_if() {
    let out = run_main(
        "if (true) if (false) System.out.println(\"inner\"); else System.out.println(\"dangling\");",
    );
    assert_eq!(out, vec!["dangling"]);
}

#[test]
fn dangling_else_skips_when_inner_if_true_without_else() {
    let out = run_main(
        "if (true) if (true) System.out.println(\"inner\"); else System.out.println(\"dangling\");",
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn dangling_else_with_outer_false_never_runs_inner() {
    let out = run_main(
        "if (false) if (true) System.out.println(\"inner\"); else System.out.println(\"dangling\"); System.out.println(\"after\");",
    );
    assert_eq!(out, vec!["after"]);
}

#[test]
fn dangling_else_on_three_level_ladder_picks_innermost_else() {
    let out = run_main(
        "int x = 1; if (x > 0) if (x < 0) System.out.println(\"a\"); else System.out.println(\"b\"); else System.out.println(\"c\");",
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn explicit_braces_prevent_dangling_else_misread() {
    let out = run_main(
        "int x = 1; if (x > 0) { if (x < 0) System.out.println(\"a\"); } else { System.out.println(\"b\"); }",
    );
    assert_eq!(out, vec!["b"]);
}

#[test]
fn triple_nested_if_all_conditions_true() {
    let out = run_main(
        "int a = 3; int b = 2; int c = 1; if (a > 0) { if (b > 0) { if (c > 0) { System.out.println(\"deep\"); } } }",
    );
    assert_eq!(out, vec!["deep"]);
}

#[test]
fn triple_nested_if_middle_level_false() {
    let out = run_main(
        "int a = 3; int b = -1; int c = 1; if (a > 0) { if (b > 0) { if (c > 0) { System.out.println(\"deep\"); } else { System.out.println(\"c-fail\"); } } else { System.out.println(\"b-fail\"); } }",
    );
    assert_eq!(out, vec!["b-fail"]);
}

#[test]
fn quadruple_nested_if_reaches_innermost_branch() {
    let out = run_main(
        "int w = 1; int x = 1; int y = 1; int z = 1; if (w == 1) if (x == 1) if (y == 1) if (z == 1) System.out.println(\"four\");",
    );
    assert_eq!(out, vec!["four"]);
}

#[test]
fn quadruple_nested_if_stops_at_second_level() {
    let out = run_main(
        "int w = 1; int x = 0; if (w == 1) if (x == 1) if (true) System.out.println(\"three\"); else System.out.println(\"two\");",
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn nested_if_else_ladder_selects_mid_level() {
    let out = run_main(
        "int n = 15; if (n < 10) { System.out.println(\"low\"); } else { if (n < 20) { System.out.println(\"mid\"); } else { if (n < 30) { System.out.println(\"high\"); } else { System.out.println(\"top\"); } } }",
    );
    assert_eq!(out, vec!["mid"]);
}

#[test]
fn nested_if_else_ladder_reaches_deepest_else() {
    let out = run_main(
        "int n = 40; if (n < 10) { System.out.println(\"low\"); } else { if (n < 20) { System.out.println(\"mid\"); } else { if (n < 30) { System.out.println(\"high\"); } else { System.out.println(\"top\"); } } }",
    );
    assert_eq!(out, vec!["top"]);
}

#[test]
fn nested_if_with_shared_outer_else() {
    let out = run_main(
        "int a = 0; int b = 5; if (a > 0) { if (b > 0) System.out.println(\"in\"); } else { System.out.println(\"out\"); }",
    );
    assert_eq!(out, vec!["out"]);
}

#[test]
fn deeply_nested_mixed_if_else_without_braces_on_inner() {
    let out = run_main(
        "int p = 2; int q = 3; if (p < 5) if (q > 5) System.out.println(\"big\"); else System.out.println(\"small\");",
    );
    assert_eq!(out, vec!["small"]);
}

#[test]
fn assignment_in_if_condition_assigns_before_compare() {
    let out = run_main("int x = 0; if ((x = 7) > 5) { System.out.println(x); }");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn assignment_in_if_condition_false_branch_skips_body() {
    let out = run_main(
        "int x = 0; if ((x = 2) > 5) { System.out.println(\"yes\"); } else { System.out.println(x); }",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn assignment_in_nested_if_condition_propagates_value() {
    let out = run_main("int x = 0; if (true) { if ((x = 4) == 4) { System.out.println(x); } }");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn logical_and_short_circuits_right_operand_on_false_left() {
    let out = run_main(
        "int c = 0; boolean ok = (c++ == 5) && (c++ == 6); System.out.println(c); System.out.println(ok);",
    );
    assert_eq!(out, vec!["1", "false"]);
}

#[test]
fn logical_and_evaluates_right_when_left_true() {
    let out = run_main(
        "int c = 0; boolean ok = (c++ == 0) && (c++ == 1); System.out.println(c); System.out.println(ok);",
    );
    assert_eq!(out, vec!["2", "true"]);
}

#[test]
fn logical_or_short_circuits_right_operand_on_true_left() {
    let out = run_main(
        "int c = 0; boolean ok = (c++ == 0) || (c++ == 99); System.out.println(c); System.out.println(ok);",
    );
    assert_eq!(out, vec!["1", "true"]);
}

#[test]
fn logical_or_evaluates_right_when_left_false() {
    let out = run_main(
        "int c = 0; boolean ok = (c++ == 5) || (c++ == 0); System.out.println(c); System.out.println(ok);",
    );
    assert_eq!(out, vec!["2", "true"]);
}

#[test]
fn side_effects_in_and_condition_inside_if_body() {
    let out = run_main(
        "int a = 0; int b = 0; if ((a++ == 0) && (b++ == 0)) { System.out.println(a + b); }",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn side_effects_in_or_condition_inside_if_body() {
    let out = run_main(
        "int a = 1; int b = 0; if ((a++ == 2) || (b++ == 0)) { System.out.println(a); System.out.println(b); }",
    );
    assert_eq!(out, vec!["2", "0"]);
}

#[test]
fn compound_and_with_three_relational_checks() {
    let out = run_main(
        "int x = 5; int y = 10; int z = 15; if (x < y && y < z && z > x) { System.out.println(\"chain\"); }",
    );
    assert_eq!(out, vec!["chain"]);
}

#[test]
fn compound_and_fails_when_middle_operand_false() {
    let out = run_main(
        "int x = 5; int y = 3; int z = 15; if (x < y && y < z && z > x) { System.out.println(\"chain\"); } else { System.out.println(\"break\"); }",
    );
    assert_eq!(out, vec!["break"]);
}

#[test]
fn compound_or_with_three_range_escapes() {
    let out =
        run_main("int v = -1; if (v < 0 || v == 5 || v > 100) { System.out.println(\"hit\"); }");
    assert_eq!(out, vec!["hit"]);
}

#[test]
fn compound_or_all_false_falls_to_else() {
    let out = run_main(
        "int v = 7; if (v < 0 || v == 5 || v > 100) { System.out.println(\"hit\"); } else { System.out.println(\"miss\"); }",
    );
    assert_eq!(out, vec!["miss"]);
}

#[test]
fn mixing_and_or_with_parentheses_groups_or_first() {
    let out = run_main(
        "boolean a = false; boolean b = true; boolean c = false; if (a || (b && c)) { System.out.println(\"yes\"); } else { System.out.println(\"no\"); }",
    );
    assert_eq!(out, vec!["no"]);
}

#[test]
fn mixing_and_or_without_parens_and_binds_tighter() {
    let out = run_main(
        "boolean a = false; boolean b = true; boolean c = true; if (a || b && c) { System.out.println(\"yes\"); } else { System.out.println(\"no\"); }",
    );
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn nested_if_with_not_and_compound_condition() {
    let out =
        run_main("int x = 3; if (!(x < 0) && (x % 2 == 1)) { System.out.println(\"odd-pos\"); }");
    assert_eq!(out, vec!["odd-pos"]);
}

#[test]
fn nested_if_else_if_inside_outer_else_branch() {
    let out = run_main(
        "int code = 2; if (code == 1) { System.out.println(\"one\"); } else { if (code == 2) { System.out.println(\"two\"); } else { System.out.println(\"other\"); } }",
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn deeply_nested_if_with_increment_in_condition() {
    let out = run_main("int n = 0; if (n++ == 0) { if (n++ == 1) { System.out.println(n); } }");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn nested_if_both_branches_use_compound_or() {
    let out = run_main(
        "int s = 90; if (s < 50 || s > 80) { if (s > 85 || s < 10) { System.out.println(\"pass\"); } }",
    );
    assert_eq!(out, vec!["pass"]);
}

#[test]
fn if_with_assignment_and_equality_in_same_condition() {
    let out = run_main(
        "int v = 0; int target = 8; if ((v = target) == 8) { System.out.println(\"eq\"); }",
    );
    assert_eq!(out, vec!["eq"]);
}

#[test]
fn nested_dangling_else_with_integer_comparison() {
    let out = run_main(
        "int a = 10; int b = 5; if (a > b) if (a < b) System.out.println(\"impossible\"); else System.out.println(\"expected\");",
    );
    assert_eq!(out, vec!["expected"]);
}

#[test]
fn four_level_nested_if_else_if_chain() {
    let out = run_main(
        "int tier = 3; if (tier == 1) { System.out.println(\"t1\"); } else { if (tier == 2) { System.out.println(\"t2\"); } else { if (tier == 3) { System.out.println(\"t3\"); } else { System.out.println(\"tx\"); } } }",
    );
    assert_eq!(out, vec!["t3"]);
}

#[test]
fn boolean_negation_with_and_in_nested_if() {
    let out = run_main(
        "boolean ready = false; boolean armed = true; if (!ready) { if (armed && !ready) { System.out.println(\"standby\"); } }",
    );
    assert_eq!(out, vec!["standby"]);
}

#[test]
fn nested_if_using_postfix_increment_in_condition() {
    let out = run_main("int i = 4; if (i++ == 4) { if (i == 5) { System.out.println(i); } }");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn side_effect_or_inside_nested_else_if() {
    let out = run_main(
        "int hits = 0; int x = 0; if (x > 0) { System.out.println(\"pos\"); } else if ((hits++ == 0) || (hits++ == 2)) { System.out.println(hits); }",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn deeply_nested_if_with_string_length_and_and() {
    let out = run_main(
        r#"String s = "java"; if (s.length() > 2) { if (s.charAt(0) == 'j' && s.endsWith("a")) { System.out.println("ok"); } }"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn nested_if_else_with_zero_and_compound_and() {
    let out = run_main(
        "int a = 0; int b = 0; if (a == 0) { if (b == 0 && a + b == 0) { System.out.println(\"zero\"); } else { System.out.println(\"other\"); } }",
    );
    assert_eq!(out, vec!["zero"]);
}
