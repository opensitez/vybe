//! Boxing, unboxing, explicit/implicit casts, and safe cast operators.
use super::helpers::run_csharp;

#[test]
fn boxing_int_to_object_wraps_value() {
    assert_eq!(
        run_csharp(r#"object boxed = 42; Console.WriteLine(boxed);"#),
        &["42"]
    );
}

#[test]
fn unboxing_casts_object_back_to_int() {
    assert_eq!(
        run_csharp(r#"object boxed = 42; int n = (int)boxed; Console.WriteLine(n);"#),
        &["42"]
    );
}

#[test]
fn unboxing_to_wrong_type_throws_invalid_cast_exception() {
    assert_eq!(
        run_csharp(
            r#"object boxed = 42;
string result = "";
try { string s = (string)boxed; }
catch(System.InvalidCastException) { result = "bad"; }
Console.WriteLine(result);"#
        ),
        &["bad"]
    );
}

#[test]
fn implicit_numeric_widening_int_to_long() {
    assert_eq!(
        run_csharp(r#"int x = 100; long y = x; Console.WriteLine(y);"#),
        &["100"]
    );
}

#[test]
fn explicit_narrowing_long_to_int_truncates() {
    assert_eq!(
        run_csharp(r#"long x = 5L; int y = (int)x; Console.WriteLine(y);"#),
        &["5"]
    );
}

#[test]
fn as_operator_returns_null_when_cast_incompatible() {
    assert_eq!(
        run_csharp(r#"object o = 42; string s = o as string; Console.WriteLine(s == null);"#),
        &["True"]
    );
}

#[test]
fn as_operator_returns_value_when_cast_compatible() {
    assert_eq!(
        run_csharp(r#"object o = "hello"; string s = o as string; Console.WriteLine(s);"#),
        &["hello"]
    );
}

#[test]
fn is_operator_with_pattern_binds_typed_variable() {
    assert_eq!(
        run_csharp(
            r#"object o = "world";
if(o is string s) Console.WriteLine(s.Length);"#
        ),
        &["5"]
    );
}

#[test]
fn double_to_int_truncates_fractional_part() {
    assert_eq!(
        run_csharp(r#"double d = 9.9; int n = (int)d; Console.WriteLine(n);"#),
        &["9"]
    );
}

#[test]
fn int_to_double_implicit_widens_without_loss() {
    assert_eq!(
        run_csharp(r#"int i = 5; double d = i; Console.WriteLine(d);"#),
        &["5"]
    );
}
