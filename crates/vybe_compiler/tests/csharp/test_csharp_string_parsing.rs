//! Parsing and converting string representations of primitive types.
use super::helpers::run_csharp;

#[test]
fn int_parse_converts_decimal_string() {
    assert_eq!(run_csharp(r#"Console.WriteLine(int.Parse("42"));"#), &["42"]);
}

#[test]
fn double_parse_with_invariant_culture() {
    assert_eq!(
        run_csharp(r#"var d=double.Parse("3.14",System.Globalization.CultureInfo.InvariantCulture);
Console.WriteLine(d);"#),
        &["3.14"]
    );
}

#[test]
fn int_try_parse_returns_false_for_non_numeric() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(int.TryParse("abc",out _));"#),
        &["False"]
    );
}

#[test]
fn bool_parse_converts_true_string() {
    assert_eq!(run_csharp(r#"Console.WriteLine(bool.Parse("True"));"#), &["True"]);
}

#[test]
fn datetime_try_parse_exact_with_format() {
    assert_eq!(
        run_csharp(r#"bool ok=System.DateTime.TryParseExact("2024-01-15","yyyy-MM-dd",
    System.Globalization.CultureInfo.InvariantCulture,
    System.Globalization.DateTimeStyles.None,out var dt);
Console.WriteLine(ok); Console.WriteLine(dt.Day);"#),
        &["True", "15"]
    );
}

#[test]
fn decimal_parse_preserves_exact_fraction() {
    assert_eq!(
        run_csharp(r#"var d=decimal.Parse("0.1",System.Globalization.CultureInfo.InvariantCulture);
Console.WriteLine(d+0.2m==0.3m);"#),
        &["True"]
    );
}

#[test]
fn enum_try_parse_succeeds_on_valid_name() {
    assert_eq!(
        run_csharp(r#"enum Color{Red,Green,Blue}
Console.WriteLine(System.Enum.TryParse<Color>("Green",out var c));
Console.WriteLine(c);"#),
        &["True", "Green"]
    );
}

#[test]
fn guid_try_parse_recognises_standard_format() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Guid.TryParse("550e8400-e29b-41d4-a716-446655440000",out _));"#),
        &["True"]
    );
}
