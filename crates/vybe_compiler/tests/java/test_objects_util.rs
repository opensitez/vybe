use crate::helpers::{run_in_main, run_main};

#[test]
fn objects_equals_both_null_returns_true() {
    let out = run_main(
        "String a = null; String b = null; System.out.println(java.util.Objects.equals(a, b));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_equals_one_null_returns_false() {
    let out = run_main(
        r#"String a = "x"; String b = null; System.out.println(java.util.Objects.equals(a, b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_equals_same_string_content_returns_true() {
    let out = run_main(
        r#"System.out.println(java.util.Objects.equals("java", "java"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_equals_different_strings_returns_false() {
    let out = run_main(
        r#"System.out.println(java.util.Objects.equals("a", "b"));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_equals_symmetric_for_integers() {
    let out = run_main(
        "Integer a = 7; Integer b = 7; System.out.println(java.util.Objects.equals(a, b)); System.out.println(java.util.Objects.equals(b, a));",
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn objects_equals_self_reference_returns_true() {
    let out = run_main(
        r#"String s = "self"; System.out.println(java.util.Objects.equals(s, s));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_hash_single_argument_nonzero() {
    let out = run_main(
        r#"System.out.println(java.util.Objects.hash("solo") != 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_hash_two_values_stable() {
    let out = run_main(
        r#"int h1 = java.util.Objects.hash("a", "b"); int h2 = java.util.Objects.hash("a", "b"); System.out.println(h1 == h2);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_hash_order_matters_for_pair() {
    let out = run_main(
        r#"int ab = java.util.Objects.hash("a", "b"); int ba = java.util.Objects.hash("b", "a"); System.out.println(ab == ba);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_hash_null_argument_allowed() {
    let out = run_main(
        "System.out.println(java.util.Objects.hash((Object) null));",
    );
    assert_eq!(out.len(), 1);
}

#[test]
fn objects_hash_three_fields_differs_from_two() {
    let out = run_main(
        r#"int two = java.util.Objects.hash("x", "y"); int three = java.util.Objects.hash("x", "y", "z"); System.out.println(two == three);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_require_non_null_returns_argument() {
    let out = run_main(
        r#"String s = java.util.Objects.requireNonNull("safe"); System.out.println(s);"#,
    );
    assert_eq!(out, vec!["safe"]);
}

#[test]
fn objects_require_non_null_with_message_returns_value() {
    let out = run_main(
        r#"String s = java.util.Objects.requireNonNull("ok", "must not be null"); System.out.println(s);"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn objects_require_non_null_throws_on_null() {
    let out = run_in_main(
        "try { java.util.Objects.requireNonNull(null); System.out.println(\"ok\"); } catch (NullPointerException e) { System.out.println(\"npe\"); }",
        "",
    );
    assert_eq!(out, vec!["npe"]);
}

#[test]
fn objects_require_non_null_message_included_in_exception() {
    let out = run_in_main(
        r#"try { java.util.Objects.requireNonNull(null, "missing"); System.out.println("ok"); } catch (NullPointerException e) { System.out.println(e.getMessage()); }"#,
        "",
    );
    assert_eq!(out, vec!["missing"]);
}

#[test]
fn objects_is_null_true_for_null() {
    let out = run_main(
        "String s = null; System.out.println(java.util.Objects.isNull(s));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_is_null_false_for_value() {
    let out = run_main(
        r#"String s = "x"; System.out.println(java.util.Objects.isNull(s));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_non_null_false_for_null() {
    let out = run_main(
        "String s = null; System.out.println(java.util.Objects.nonNull(s));",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_non_null_true_for_value() {
    let out = run_main(
        r#"String s = "x"; System.out.println(java.util.Objects.nonNull(s));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_compare_ascending_integers() {
    let out = run_main(
        "System.out.println(java.util.Objects.compare(1, 3, (a, b) -> a - b));",
    );
    assert_eq!(out, vec!["-2"]);
}

#[test]
fn objects_compare_descending_integers() {
    let out = run_main(
        "System.out.println(java.util.Objects.compare(5, 2, (a, b) -> b - a));",
    );
    assert_eq!(out, vec!["-3"]);
}

#[test]
fn objects_compare_equal_values_zero() {
    let out = run_main(
        r#"System.out.println(java.util.Objects.compare("a", "a", (x, y) -> x.compareTo(y)));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn objects_compare_first_null_negative_without_comparator_call() {
    let out = run_main(
        "System.out.println(java.util.Objects.compare(null, \"z\", (a, b) -> 0));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn objects_compare_second_null_positive_without_comparator_call() {
    let out = run_main(
        r#"System.out.println(java.util.Objects.compare("z", null, (a, b) -> 0));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn objects_compare_both_null_zero() {
    let out = run_main(
        "System.out.println(java.util.Objects.compare(null, null, (a, b) -> 0));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn objects_to_string_fallback_for_null() {
    let out = run_main(
        r#"String s = null; System.out.println(java.util.Objects.toString(s, "fallback"));"#,
    );
    assert_eq!(out, vec!["fallback"]);
}

#[test]
fn objects_to_string_uses_value_when_non_null() {
    let out = run_main(
        r#"String s = "live"; System.out.println(java.util.Objects.toString(s, "fallback"));"#,
    );
    assert_eq!(out, vec!["live"]);
}

#[test]
fn objects_to_string_single_arg_non_null() {
    let out = run_main(
        r#"System.out.println(java.util.Objects.toString("plain"));"#,
    );
    assert_eq!(out, vec!["plain"]);
}

#[test]
fn objects_to_string_single_arg_null_prints_null() {
    let out = run_main(
        "String s = null; System.out.println(java.util.Objects.toString(s));",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn objects_equals_integer_and_string_value_of() {
    let out = run_main(
        "Integer n = 42; System.out.println(java.util.Objects.equals(n, Integer.valueOf(42)));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_hash_integer_wrapper_matches_repeat() {
    let out = run_main(
        "int h1 = java.util.Objects.hash(Integer.valueOf(1)); int h2 = java.util.Objects.hash(Integer.valueOf(1)); System.out.println(h1 == h2);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_equals_empty_strings() {
    let out = run_main(
        r#"System.out.println(java.util.Objects.equals("", ""));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_is_null_and_non_null_are_opposites() {
    let out = run_main(
        r#"String s = "data"; System.out.println(java.util.Objects.isNull(s)); System.out.println(java.util.Objects.nonNull(s));"#,
    );
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn objects_require_non_null_integer_wrapper() {
    let out = run_main(
        "Integer n = java.util.Objects.requireNonNull(Integer.valueOf(8)); System.out.println(n);",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn objects_compare_strings_lexicographic() {
    let out = run_main(
        r#"System.out.println(java.util.Objects.compare("apple", "banana", (a, b) -> a.compareTo(b)));"#,
    );
    assert_eq!(out.len(), 1);
    assert!(out[0].starts_with('-'));
}

#[test]
fn objects_to_string_fallback_empty_string() {
    let out = run_main(
        r#"System.out.println(java.util.Objects.toString(null, ""));"#,
    );
    assert_eq!(out, vec![""]);
}

#[test]
fn objects_hash_mixed_null_and_value() {
    let out = run_main(
        r#"int h = java.util.Objects.hash(null, "x"); System.out.println(h == java.util.Objects.hash(null, "x"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_equals_after_concatenation() {
    let out = run_main(
        r#"String built = "ja" + "va"; System.out.println(java.util.Objects.equals(built, "java"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_to_string_on_integer_calls_value_of() {
    let out = run_main(
        "System.out.println(java.util.Objects.toString(99));",
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn objects_equals_boolean_boxed_values() {
    let out = run_main(
        "System.out.println(java.util.Objects.equals(Boolean.TRUE, Boolean.TRUE));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_compare_with_natural_integer_order() {
    let out = run_main(
        "System.out.println(java.util.Objects.compare(10, 10, Integer::compare));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn objects_hash_code_differs_for_distinct_strings() {
    let out = run_main(
        r#"System.out.println(java.util.Objects.hash("one") == java.util.Objects.hash("two"));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_non_null_used_in_conditional_branch() {
    let out = run_main(
        r#"String s = "go"; if (java.util.Objects.nonNull(s)) { System.out.println(s); }"#,
    );
    assert_eq!(out, vec!["go"]);
}

#[test]
fn objects_is_null_used_in_conditional_branch() {
    let out = run_main(
        "String s = null; if (java.util.Objects.isNull(s)) { System.out.println(\"missing\"); }",
    );
    assert_eq!(out, vec!["missing"]);
}

#[test]
fn objects_to_string_fallback_on_integer_null_reference() {
    let out = run_main(
        "Integer n = null; System.out.println(java.util.Objects.toString(n, \"none\"));",
    );
    assert_eq!(out, vec!["none"]);
}