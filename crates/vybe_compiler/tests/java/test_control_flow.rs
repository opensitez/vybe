use crate::helpers::run_main;

#[test]
fn if_without_else_skips_false_branch() {
    let out = run_main(
        "int x = 1; if (x > 5) { System.out.println(\"yes\"); } System.out.println(\"done\");",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn if_else_picks_else_when_condition_false() {
    let out = run_main(
        "int x = 2; if (x > 5) { System.out.println(\"big\"); } else { System.out.println(\"small\"); }",
    );
    assert_eq!(out, vec!["small"]);
}

#[test]
fn else_if_chain_selects_first_match() {
    let out = run_main(
        "int score = 75; if (score >= 90) { System.out.println(\"A\"); } else if (score >= 70) { System.out.println(\"B\"); } else { System.out.println(\"C\"); }",
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn ternary_expression_selects_false_branch() {
    let out = run_main(
        "int n = 4; String label = (n % 2 == 0) ? \"even\" : \"odd\"; System.out.println(label);",
    );
    assert_eq!(out, vec!["even"]);
}

#[test]
fn logical_or_short_circuits_when_first_true() {
    let out = run_main(
        "boolean called = false; boolean result = true || (called = true); System.out.println(result); System.out.println(called);",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn logical_and_short_circuits_when_first_false() {
    let out = run_main(
        "boolean called = false; boolean result = false && (called = true); System.out.println(result); System.out.println(called);",
    );
    assert_eq!(out, vec!["false", "false"]);
}
