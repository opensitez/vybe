//! `System.Convert` class: cross-type conversions and `ChangeType`.
use super::helpers::run_csharp;

#[test]
fn convert_to_int32_from_string() {
    assert_eq!(
        run_csharp(
            r#"// conversion_methods
Console.WriteLine(System.Convert.ToInt32("99"));"#
        ),
        &["99"]
    );
}

#[test]
fn convert_to_double_from_int() {
    assert_eq!(
        run_csharp(
            r#"double d=System.Convert.ToDouble(7);
Console.WriteLine(d);"#
        ),
        &["7"]
    );
}

#[test]
fn convert_to_string_from_bool() {
    assert_eq!(
        run_csharp(
            r#"// conversion_methods
Console.WriteLine(System.Convert.ToString(true));"#
        ),
        &["True"]
    );
}

#[test]
fn convert_to_boolean_from_int_one_is_true() {
    assert_eq!(
        run_csharp(
            r#"// conversion_methods
Console.WriteLine(System.Convert.ToBoolean(1));"#
        ),
        &["True"]
    );
}

#[test]
fn convert_to_boolean_from_zero_is_false() {
    assert_eq!(
        run_csharp(
            r#"// conversion_methods
Console.WriteLine(System.Convert.ToBoolean(0));"#
        ),
        &["False"]
    );
}

#[test]
fn convert_to_char_from_int_gives_unicode_char() {
    assert_eq!(
        run_csharp(
            r#"// conversion_methods
Console.WriteLine(System.Convert.ToChar(65));"#
        ),
        &["A"]
    );
}

#[test]
fn convert_change_type_dynamically_converts_to_target_type() {
    assert_eq!(
        run_csharp(
            r#"object result=System.Convert.ChangeType("42",typeof(int));
Console.WriteLine(result);"#
        ),
        &["42"]
    );
}
