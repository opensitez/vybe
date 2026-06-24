use crate::helpers::run_main;

#[test]
fn int_to_long_widening_preserves_value() {
    let out = run_main("int n = 42; long wide = n; System.out.println(wide);");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn double_to_int_truncates_fraction() {
    let out = run_main("double d = 9.9; int n = (int) d; System.out.println(n);");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn string_to_integer_via_parse_int() {
    let out = run_main("int n = Integer.parseInt(\"123\"); System.out.println(n);");
    assert_eq!(out, vec!["123"]);
}

#[test]
fn autoboxing_wraps_primitive_in_wrapper() {
    let out = run_main("Integer boxed = 7; System.out.println(boxed);");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn unboxing_extracts_primitive_from_wrapper() {
    let out = run_main("Integer boxed = 12; int n = boxed; System.out.println(n + 1);");
    assert_eq!(out, vec!["13"]);
}

#[test]
fn instanceof_before_downcast_to_string() {
    let out = run_main(
        "Object o = \"java\"; if (o instanceof String s) { System.out.println(s.toUpperCase()); }",
    );
    assert_eq!(out, vec!["JAVA"]);
}
