//! `Guid` parse, format, and equality semantics.
use super::helpers::run_csharp;

#[test]
fn guid_empty_has_all_zero_bytes() {
    assert_eq!(
        run_csharp(
            r#"
var id = System.Guid.Empty;
Console.WriteLine(id == new System.Guid("00000000-0000-0000-0000-000000000000"));
"#
        ),
        &["True"]
    );
}

#[test]
fn guid_parse_accepts_standard_hyphenated_representation() {
    assert_eq!(
        run_csharp(
            r#"
var id = System.Guid.Parse("11111111-2222-3333-4444-555555555555");
Console.WriteLine(id.ToString().StartsWith("11111111"));
"#
        ),
        &["True"]
    );
}

#[test]
fn guid_try_parse_returns_false_for_invalid_literal() {
    assert_eq!(
        run_csharp(
            r#"
System.Guid value;
var ok = System.Guid.TryParse("not-a-guid", out value);
Console.WriteLine(ok);
"#
        ),
        &["False"]
    );
}

#[test]
fn guid_to_string_with_format_specifier_renders_hyphenated_value() {
    assert_eq!(
        run_csharp(
            r#"
var id = System.Guid.Parse("11111111-2222-3333-4444-555555555555");
Console.WriteLine(id.ToString("D").StartsWith("11111111"));
"#
        ),
        &["True"]
    );
}

#[test]
fn guid_new_guid_produces_unique_values_on_successive_calls() {
    assert_eq!(
        run_csharp(
            r#"
var left = System.Guid.NewGuid();
var right = System.Guid.NewGuid();
Console.WriteLine(left != right);
"#
        ),
        &["True"]
    );
}
