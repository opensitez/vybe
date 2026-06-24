use crate::helpers::run_main;

#[test]
fn int_equality_true_for_same_value() {
    let out = run_main("System.out.println(5 == 5);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_equality_false_for_different_values() {
    let out = run_main("System.out.println(5 == 6);");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn int_inequality_true_for_different_values() {
    let out = run_main("System.out.println(5 != 6);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_less_than_orders_smaller_first() {
    let out = run_main("System.out.println(2 < 3);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_less_than_or_equal_accepts_equal_values() {
    let out = run_main("System.out.println(3 <= 3);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn int_greater_than_orders_larger_first() {
    let out = run_main("System.out.println(4 > 5);");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn int_greater_than_or_equal_accepts_equal_values() {
    let out = run_main("System.out.println(7 >= 7);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn long_equality_compares_wide_values() {
    let out = run_main("System.out.println(1_000_000L == 1_000_000L);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn long_less_than_orders_smaller_wide_value() {
    let out = run_main("System.out.println(5L < 9L);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_equality_matches_same_fraction() {
    let out = run_main("System.out.println(2.5 == 2.5);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_less_than_orders_smaller_fraction() {
    let out = run_main("System.out.println(1.1 < 2.2);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn double_nan_not_equal_to_itself() {
    let out = run_main("System.out.println(Double.NaN == Double.NaN);");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn double_is_nan_rejects_equality_with_finite_value() {
    let out = run_main("System.out.println(Double.NaN == 1.0);");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn boolean_equality_matches_same_literal() {
    let out = run_main("System.out.println(true == true);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn boolean_inequality_detects_opposite_literals() {
    let out = run_main("System.out.println(true != false);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn char_less_than_orders_smaller_code_unit() {
    let out = run_main("System.out.println('a' < 'b');");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn char_equality_matches_same_literal() {
    let out = run_main("System.out.println('X' == 'X');");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_literal_equality_compares_same_text() {
    let out = run_main(r#"System.out.println("java" == "java");"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_variable_equality_true_for_same_literal_text() {
    let out = run_main(r#"String a = "java"; String b = "java"; System.out.println(a == b);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_equals_compares_matching_content() {
    let out = run_main(
        r#"String a = "java"; String b = "java"; System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_equals_rejects_different_content() {
    let out = run_main(
        r#"String a = "java"; String b = "kotlin"; System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn string_compare_to_negative_when_lexicographically_less() {
    let out = run_main(r#"System.out.println("apple".compareTo("banana"));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn string_compare_to_zero_for_identical_text() {
    let out = run_main(r#"System.out.println("same".compareTo("same"));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn string_compare_to_positive_when_lexicographically_greater() {
    let out = run_main(r#"System.out.println("zebra".compareTo("apple"));"#);
    assert_eq!(out, vec!["25"]);
}

#[test]
fn string_compare_to_ignore_case_treats_ascii_case_as_equal() {
    let out = run_main(r#"System.out.println("ABC".compareToIgnoreCase("abc"));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn null_equality_detects_two_null_references() {
    let out = run_main("String a = null; String b = null; System.out.println(a == b);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn null_inequality_detects_non_null_value() {
    let out = run_main(r#"String a = null; String b = "x"; System.out.println(a != b);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn boxed_integer_equality_compares_unboxed_values() {
    let out = run_main("Integer a = 7; Integer b = 7; System.out.println(a == b);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn integer_compare_orders_smaller_first() {
    let out = run_main(
        "System.out.println(Integer.compare(5, 8)); System.out.println(Integer.compare(8, 5));",
    );
    assert_eq!(out, vec!["-1", "1"]);
}

#[test]
fn chained_int_comparison_in_range_check() {
    let out = run_main("int n = 5; System.out.println(n >= 1 && n <= 10);");
    assert_eq!(out, vec!["true"]);
}
