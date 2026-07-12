use crate::helpers::{run_in_main, run_main};

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
fn non_null_reference_is_not_equal_to_null() {
    let out = run_main(
        r#"String s = "data"; System.out.println(s == null); System.out.println(s != null);"#,
    );
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn null_safe_branch_skips_method_call_on_null() {
    let out = run_main(
        "String s = null; if (s == null) { System.out.println(\"missing\"); } else { System.out.println(s.length()); }",
    );
    assert_eq!(out, vec!["missing"]);
}

#[test]
fn non_null_branch_invokes_method_when_present() {
    let out = run_main(
        r#"String s = "abc"; if (s == null) { System.out.println("missing"); } else { System.out.println(s.length()); }"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn instance_equals_returns_false_for_null_argument() {
    let out = run_main(r#"String s = "x"; System.out.println(s.equals(null));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instance_equals_returns_true_for_equal_non_null() {
    let out = run_main(r#"String a = "vybe"; String b = "vybe"; System.out.println(a.equals(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_equals_handles_two_null_references() {
    let out = run_main(
        "String a = null; String b = null; System.out.println(java.util.Objects.equals(a, b));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_equals_handles_one_null_reference() {
    let out = run_main(
        r#"String a = "x"; String b = null; System.out.println(java.util.Objects.equals(a, b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_is_null_true_for_null_reference() {
    let out = run_main("String s = null; System.out.println(java.util.Objects.isNull(s));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_is_null_false_for_present_reference() {
    let out = run_main(r#"String s = "ok"; System.out.println(java.util.Objects.isNull(s));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_non_null_false_for_null_reference() {
    let out = run_main("String s = null; System.out.println(java.util.Objects.nonNull(s));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_non_null_true_for_present_reference() {
    let out = run_main(r#"String s = "ok"; System.out.println(java.util.Objects.nonNull(s));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_require_non_null_returns_argument_when_present() {
    let out =
        run_main(r#"String s = java.util.Objects.requireNonNull("safe"); System.out.println(s);"#);
    assert_eq!(out, vec!["safe"]);
}

#[test]
fn optional_empty_is_not_present() {
    let out = run_main(
        "java.util.Optional<String> opt = java.util.Optional.empty(); System.out.println(opt.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_of_wraps_non_null_value() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.of("value"); System.out.println(opt.isPresent());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_of_nullable_empty_when_null() {
    let out = run_main(
        "String s = null; java.util.Optional<String> opt = java.util.Optional.ofNullable(s); System.out.println(opt.isPresent());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn optional_of_nullable_present_when_non_null() {
    let out = run_main(
        r#"String s = "live"; java.util.Optional<String> opt = java.util.Optional.ofNullable(s); System.out.println(opt.isPresent());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn optional_or_else_supplies_default_for_empty() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.empty(); System.out.println(opt.orElse("fallback"));"#,
    );
    assert_eq!(out, vec!["fallback"]);
}

#[test]
fn optional_or_else_returns_value_when_present() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.of("real"); System.out.println(opt.orElse("fallback"));"#,
    );
    assert_eq!(out, vec!["real"]);
}

#[test]
fn ternary_null_coalesce_substitutes_default_for_null() {
    let out = run_main(
        r#"String s = null; String v = s != null ? s : "default"; System.out.println(v);"#,
    );
    assert_eq!(out, vec!["default"]);
}

#[test]
fn ternary_null_coalesce_keeps_present_value() {
    let out = run_main(
        r#"String s = "keep"; String v = s != null ? s : "default"; System.out.println(v);"#,
    );
    assert_eq!(out, vec!["keep"]);
}

#[test]
fn optional_or_else_get_supplies_lazy_default_for_empty() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.empty(); System.out.println(opt.orElseGet(() -> "lazy"));"#,
    );
    assert_eq!(out, vec!["lazy"]);
}

#[test]
fn optional_if_present_runs_consumer_when_value_exists() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.of("hi"); opt.ifPresent(v -> System.out.println(v));"#,
    );
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn null_string_concatenation_with_literal() {
    let out = run_main(r#"String s = null; System.out.println("prefix" + s);"#);
    assert_eq!(out, vec!["prefixnull"]);
}

#[test]
fn assigning_null_to_object_reference() {
    let out = run_main(r#"String s = "temp"; s = null; System.out.println(s == null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn null_integer_wrapper_reference_prints_null() {
    let out = run_main("Integer n = null; System.out.println(n);");
    assert_eq!(out, vec!["null"]);
}

#[test]
fn optional_nested_or_else_picks_inner_default() {
    let out = run_main(
        r#"java.util.Optional<String> opt = java.util.Optional.empty(); String v = opt.orElse(null); String out = v != null ? v : "final"; System.out.println(out);"#,
    );
    assert_eq!(out, vec!["final"]);
}

#[test]
fn objects_equals_distinguishes_different_non_null_strings() {
    let out = run_main(r#"System.out.println(java.util.Objects.equals("a", "b"));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn null_check_before_dereference_avoids_length_call() {
    let out = run_main(
        r#"String s = null; int len = s != null ? s.length() : 0; System.out.println(len);"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn optional_of_nullable_from_method_result() {
    let out = run_in_main(
        r#"java.util.Optional<Integer> opt = java.util.Optional.ofNullable(pickNull()); System.out.println(opt.isPresent());"#,
        r#"static Integer pickNull() { return null; }"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn require_non_null_with_message_expression_compiles_for_present_value() {
    let out = run_main(
        r#"String s = java.util.Objects.requireNonNull("ok", "must not be null"); System.out.println(s);"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn null_coalesce_via_optional_chain_on_empty() {
    let out = run_main(
        r#"String s = null; String v = java.util.Optional.ofNullable(s).orElse("via-optional"); System.out.println(v);"#,
    );
    assert_eq!(out, vec!["via-optional"]);
}

#[test]
fn null_coalesce_via_optional_chain_keeps_value() {
    let out = run_main(
        r#"String s = "direct"; String v = java.util.Optional.ofNullable(s).orElse("via-optional"); System.out.println(v);"#,
    );
    assert_eq!(out, vec!["direct"]);
}

#[test]
fn comparing_null_to_null_with_equality() {
    let out = run_main("String a = null; String b = null; System.out.println(a == b);");
    assert_eq!(out, vec!["true"]);
}
