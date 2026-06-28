use crate::helpers::{run_in_main, run_main};

#[test]
fn simple_ternary_true_branch_for_int() {
    let out = run_main("int n = 5; int v = (n > 0) ? 1 : 0; System.out.println(v);");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn simple_ternary_false_branch_for_int() {
    let out = run_main("int n = -1; int v = (n > 0) ? 1 : 0; System.out.println(v);");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn ternary_with_string_branches() {
    let out = run_main(
        r#"int n = 4; String label = (n % 2 == 0) ? "even" : "odd"; System.out.println(label);"#,
    );
    assert_eq!(out, vec!["even"]);
}

#[test]
fn nested_ternary_selects_inner_true_branch() {
    let out =
        run_main("int n = 5; int v = (n > 3) ? ((n > 4) ? 1 : 2) : 0; System.out.println(v);");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn nested_ternary_selects_inner_false_branch() {
    let out =
        run_main("int n = 4; int v = (n > 3) ? ((n > 4) ? 1 : 2) : 0; System.out.println(v);");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn ternary_with_boolean_result() {
    let out = run_main(
        "int a = 2; int b = 5; boolean ok = (a < b) ? true : false; System.out.println(ok);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn ternary_assign_to_local_variable() {
    let out = run_main("int x = 10; int y = (x > 5) ? x : 0; System.out.println(y);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn ternary_with_comparison_on_doubles() {
    let out = run_main(
        "double a = 2.5; double b = 1.1; String tag = (a > b) ? \"gt\" : \"le\"; System.out.println(tag);",
    );
    assert_eq!(out, vec!["gt"]);
}

#[test]
fn ternary_double_branches() {
    let out = run_main("double x = 3.0; double y = (x > 2.0) ? 9.5 : 1.5; System.out.println(y);");
    assert_eq!(out, vec!["9.5"]);
}

#[test]
fn ternary_char_branches() {
    let out = run_main("boolean upper = true; char c = upper ? 'A' : 'a'; System.out.println(c);");
    assert_eq!(out, vec!["A"]);
}

#[test]
fn ternary_long_branches() {
    let out = run_main("long n = 1_000_000L; long v = (n > 0L) ? 99L : 0L; System.out.println(v);");
    assert_eq!(out, vec!["99"]);
}

#[test]
fn ternary_with_null_object_branch() {
    let out = run_main(
        "boolean missing = true; String s = missing ? null : \"x\"; System.out.println(s);",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn ternary_object_string_reference_branches() {
    let out = run_main(
        r#"boolean pickFirst = false; String a = "first"; String b = "second"; String chosen = pickFirst ? a : b; System.out.println(chosen);"#,
    );
    assert_eq!(out, vec!["second"]);
}

#[test]
fn ternary_integer_wrapper_branches() {
    let out = run_main(
        "boolean useBig = true; Integer v = useBig ? Integer.valueOf(100) : Integer.valueOf(1); System.out.println(v);",
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn ternary_nested_three_levels_deep() {
    let out = run_main(
        "int n = 6; int v = (n > 5) ? ((n > 7) ? 3 : ((n > 6) ? 2 : 1)) : 0; System.out.println(v);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn ternary_embedded_directly_in_println() {
    let out = run_main("int n = 3; System.out.println((n % 2 == 0) ? \"even\" : \"odd\");");
    assert_eq!(out, vec!["odd"]);
}

#[test]
fn ternary_with_modulo_parity_condition() {
    let out = run_main(
        "int n = 8; String parity = (n % 2 == 0) ? \"even\" : \"odd\"; System.out.println(parity);",
    );
    assert_eq!(out, vec!["even"]);
}

#[test]
fn ternary_primitive_int_vs_zero_default() {
    let out =
        run_main("int score = 55; int grade = (score >= 60) ? 1 : 0; System.out.println(grade);");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn ternary_chained_in_arithmetic_expression() {
    let out = run_main(
        "int a = 2; int b = 3; int total = (a > b ? a : b) + 10; System.out.println(total);",
    );
    assert_eq!(out, vec!["13"]);
}

#[test]
fn ternary_mixed_primitive_and_boxed_integer() {
    let types = r#"
        static class Pick {
            static int choose(boolean flag) {
                Integer boxed = flag ? 7 : 0;
                return boxed;
            }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Pick.choose(true)); System.out.println(Pick.choose(false));",
        types,
    );
    assert_eq!(out, vec!["7", "0"]);
}
