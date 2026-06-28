use crate::helpers::run_main;

#[test]
fn null_reference_prints_null() {
    let out = run_main("String s = null; System.out.println(s);");
    assert_eq!(out, vec!["null"]);
}

#[test]
fn null_equality_only_matches_null() {
    let out =
        run_main("String s = null; System.out.println(s == null); System.out.println(s != null);");
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn null_safe_field_access_via_explicit_check() {
    let out = run_main(
        "String s = null; if (s == null) { System.out.println(\"missing\"); } else { System.out.println(s.length()); }",
    );
    assert_eq!(out, vec!["missing"]);
}

#[test]
fn objects_equals_handles_null_argument() {
    let out = run_main("String s = \"x\"; System.out.println(s.equals(null));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_of_nullable_empty_when_null() {
    let out = run_main(
        "String s = null; java.util.Optional<String> opt = java.util.Optional.ofNullable(s); System.out.println(opt.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_or_else_supplies_default_for_empty() {
    let out = run_main(
        "java.util.Optional<String> opt = java.util.Optional.empty(); System.out.println(opt.orElse(\"fallback\"));",
    );
    assert_eq!(out, vec!["fallback"]);
}
