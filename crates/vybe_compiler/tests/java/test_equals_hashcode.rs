use crate::helpers::{run_in_main, run_main};

#[test]
fn string_equals_same_content_returns_true() {
    let out = run_main(
        r#"String a = "java"; String b = "java"; System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_equals_different_content_returns_false() {
    let out = run_main(
        r#"String a = "java"; String b = "kotlin"; System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn string_equals_null_argument_returns_false() {
    let out = run_main(r#"String s = "x"; System.out.println(s.equals(null));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn string_hashcode_stable_across_repeated_calls() {
    let out = run_main(
        r#"String s = "stable"; int h1 = s.hashCode(); int h2 = s.hashCode(); System.out.println(h1 == h2);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn distinct_strings_produce_different_hashcodes() {
    let out = run_main(
        r#"String a = "alpha"; String b = "beta"; System.out.println(a.hashCode() == b.hashCode());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn string_equals_ignore_case_for_ascii_pairs() {
    let out = run_main(
        r#"String a = "Java"; String b = "java"; System.out.println(a.equalsIgnoreCase(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn integer_wrapper_equals_same_value() {
    let out = run_main("Integer a = 7; Integer b = 7; System.out.println(a.equals(b));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn integer_wrapper_not_equal_different_values() {
    let out = run_main("Integer a = 7; Integer b = 8; System.out.println(a.equals(b));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_is_null_true_for_null_reference() {
    let out = run_main("String s = null; System.out.println(java.util.Objects.isNull(s));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_is_null_false_for_non_null_reference() {
    let out = run_main(
        r#"String s = "x"; System.out.println(java.util.Objects.isNull(s));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_non_null_false_for_null_reference() {
    let out = run_main("String s = null; System.out.println(java.util.Objects.nonNull(s));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn objects_non_null_true_for_non_null_reference() {
    let out = run_main(
        r#"String s = "x"; System.out.println(java.util.Objects.nonNull(s));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn null_reference_equality_only_matches_null() {
    let out = run_main(
        "String a = null; String b = null; System.out.println(a == b); System.out.println(a != null);",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn string_literal_deep_equality_against_variable() {
    let out = run_main(
        r#"String built = "ja" + "va"; System.out.println(built.equals("java"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_equals_self_reference_returns_true() {
    let out = run_main(r#"String s = "same"; System.out.println(s.equals(s));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_equals_empty_strings_match() {
    let out = run_main(
        r#"String a = ""; String b = ""; System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn custom_equals_method_compares_field_values() {
    let types = r#"
        static class Pair {
            int x;
            int y;
            Pair(int x, int y) { this.x = x; this.y = y; }
            boolean equals(Pair other) {
                return other != null && other.x == x && other.y == y;
            }
        }
    "#;
    let out = run_in_main(
        "Pair a = new Pair(1, 2); Pair b = new Pair(1, 2); Pair c = new Pair(1, 3); System.out.println(a.equals(b)); System.out.println(a.equals(c));",
        types,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn hashcode_differs_after_changing_distinct_strings() {
    let out = run_main(
        r#"String a = "one"; String b = "two"; System.out.println(a.hashCode() != b.hashCode());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_equals_both_null_via_helper_semantics() {
    let out = run_main(
        "String a = null; String b = null; boolean same = (a == b) || (a != null && a.equals(b)); System.out.println(same);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn objects_equals_one_null_other_value_is_false() {
    let out = run_main(
        r#"String a = null; String b = "x"; boolean same = (a == b) || (a != null && a.equals(b)); System.out.println(same);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn equals_combined_with_logical_and_short_circuits() {
    let out = run_main(
        r#"String a = "ok"; String b = "ok"; System.out.println(a != null && a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_equals_after_concatenation_builds_deep_content() {
    let out = run_main(
        r#"String left = "hel"; String right = "lo"; String whole = left + right; System.out.println(whole.equals("hello"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_not_equal_to_integer_wrapper_string_form() {
    let out = run_main(
        r#"String s = "7"; Integer n = 7; System.out.println(s.equals(n.toString()));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn immutable_string_hashcode_unchanged_after_reassignment() {
    let out = run_main(
        r#"String s = "alpha"; int h1 = s.hashCode(); s = "alpha"; int h2 = s.hashCode(); System.out.println(h1 == h2);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn deep_string_equality_with_unicode_escape_sequences() {
    let out = run_main(
        r#"String a = "\u004aava"; String b = "Java"; System.out.println(a.equalsIgnoreCase(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}
