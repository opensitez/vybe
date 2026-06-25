//! Culture-aware string operations: `CultureInfo`, ordinal vs. invariant comparisons.
use super::helpers::run_csharp;

#[test]
fn string_compare_invariant_culture_ignores_locale() {
    assert_eq!(
        run_csharp(r#"int r=string.Compare("hello","HELLO",System.StringComparison.InvariantCultureIgnoreCase);
Console.WriteLine(r==0);"#),
        &["True"]
    );
}

#[test]
fn to_upper_with_invariant_culture_uses_standard_casing() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("hello".ToUpperInvariant());"#),
        &["HELLO"]
    );
}

#[test]
fn to_lower_with_invariant_culture() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("HELLO".ToLowerInvariant());"#),
        &["hello"]
    );
}

#[test]
fn ordinal_comparison_is_byte_by_byte() {
    assert_eq!(
        run_csharp(r#"int r=string.CompareOrdinal("a","A");
Console.WriteLine(r>0);"#),
        &["True"]
    );
}

#[test]
fn string_equals_ordinal_ignores_locale_specific_rules() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("Abc".Equals("abc",System.StringComparison.OrdinalIgnoreCase));"#),
        &["True"]
    );
}

#[test]
fn invariant_culture_tostring_for_double_uses_dot_separator() {
    assert_eq!(
        run_csharp(r#"double d=1.5;
Console.WriteLine(d.ToString(System.Globalization.CultureInfo.InvariantCulture));"#),
        &["1.5"]
    );
}
