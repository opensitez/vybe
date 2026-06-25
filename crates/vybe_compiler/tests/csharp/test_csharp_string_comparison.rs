//! String comparison modes: ordinal, case-insensitive, invariant culture, `StringComparer`.
use super::helpers::run_csharp;

#[test]
fn ordinal_comparison_is_case_sensitive_by_default() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine(string.Compare("A","a",System.StringComparison.Ordinal) != 0);"#
        ),
        &["True"]
    );
}

#[test]
fn ordinal_ignore_case_treats_same_letters_as_equal() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine(string.Compare("Hello","hello",System.StringComparison.OrdinalIgnoreCase) == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn equals_with_string_comparison_case_insensitive() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine("ABC".Equals("abc",System.StringComparison.OrdinalIgnoreCase));"#
        ),
        &["True"]
    );
}

#[test]
fn contains_with_string_comparison_case_insensitive() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine("Hello World".Contains("world",System.StringComparison.OrdinalIgnoreCase));"#
        ),
        &["True"]
    );
}

#[test]
fn string_comparer_ordinal_ignore_case_used_in_sorted_set() {
    assert_eq!(
        run_csharp(
            r#"var set = new System.Collections.Generic.SortedSet<string>(
    System.StringComparer.OrdinalIgnoreCase);
set.Add("Apple"); set.Add("apple");
Console.WriteLine(set.Count);"#
        ),
        &["1"]
    );
}

#[test]
fn index_of_with_string_comparison_finds_case_insensitive() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine("fooBAR".IndexOf("bar",System.StringComparison.OrdinalIgnoreCase));"#
        ),
        &["3"]
    );
}

#[test]
fn starts_with_case_insensitive_returns_true_for_different_case_prefix() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine("HELLO".StartsWith("hell",System.StringComparison.OrdinalIgnoreCase));"#
        ),
        &["True"]
    );
}
