//! All pattern matching forms: type, const, relational, logical, tuple.
use super::helpers::run_csharp;

#[test]
fn type_pattern_matches_int_in_if_is_expression() {
    assert_eq!(
        run_csharp(
            r#"object o = 5; if(o is int n) Console.WriteLine(n); else Console.WriteLine(0);"#
        ),
        &["5"]
    );
}

#[test]
fn constant_pattern_matches_specific_literal_value() {
    assert_eq!(
        run_csharp(
            r#"int x = 3;
string result = x switch { 1 => "one", 2 => "two", 3 => "three", _ => "other" };
Console.WriteLine(result);"#
        ),
        &["three"]
    );
}

#[test]
fn relational_pattern_compares_numeric_range() {
    assert_eq!(
        run_csharp(
            r#"int score = 85;
string grade = score switch { >= 90 => "A", >= 80 => "B", >= 70 => "C", _ => "F" };
Console.WriteLine(grade);"#
        ),
        &["B"]
    );
}

#[test]
fn logical_and_pattern_requires_both_sub_patterns() {
    assert_eq!(
        run_csharp(
            r#"int n = 15;
Console.WriteLine(n is > 10 and < 20);"#
        ),
        &["True"]
    );
}

#[test]
fn logical_or_pattern_matches_if_either_sub_pattern_holds() {
    assert_eq!(
        run_csharp(
            r#"int n = 5;
Console.WriteLine(n is 3 or 5 or 7);"#
        ),
        &["True"]
    );
}

#[test]
fn not_pattern_inverts_any_sub_pattern() {
    assert_eq!(
        run_csharp(r#"object o = "text"; Console.WriteLine(o is not int);"#),
        &["True"]
    );
}

#[test]
fn tuple_pattern_deconstructs_two_element_tuple() {
    assert_eq!(
        run_csharp(
            r#"var point = (1, 0);
string axis = point switch {
    (0, 0) => "origin",
    (_, 0) => "x-axis",
    (0, _) => "y-axis",
    _       => "other"
};
Console.WriteLine(axis);"#
        ),
        &["x-axis"]
    );
}

#[test]
fn property_pattern_reads_nested_member() {
    assert_eq!(
        run_csharp(
            r#"class Rect { public int W, H; }
object r = new Rect { W=10, H=5 };
string size = r switch { Rect { W: > 8 } => "wide", _ => "narrow" };
Console.WriteLine(size);"#
        ),
        &["wide"]
    );
}

#[test]
fn switch_expression_with_null_arm() {
    assert_eq!(
        run_csharp(
            r#"string value = null;
string result = value switch { null => "nothing", var s => s };
Console.WriteLine(result);"#
        ),
        &["nothing"]
    );
}

#[test]
fn var_pattern_binds_any_value_in_switch_arm() {
    assert_eq!(
        run_csharp(
            r#"object o = 42;
string result = o switch { var x when x is int n && n > 10 => "big int", _ => "other" };
Console.WriteLine(result);"#
        ),
        &["big int"]
    );
}
