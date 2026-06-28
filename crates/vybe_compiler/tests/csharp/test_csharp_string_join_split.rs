//! `string.Join` and `string.Split` edge cases.
use super::helpers::run_csharp;

#[test]
fn join_array_with_separator() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Join(",",new[]{"a","b","c"}));"#),
        &["a,b,c"]
    );
}

#[test]
fn join_with_empty_separator_concatenates() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Join("",new[]{"a","b","c"}));"#),
        &["abc"]
    );
}

#[test]
fn split_removes_empty_entries_option() {
    assert_eq!(
        run_csharp(
            r#"var parts="a,,b,,c".Split(',',System.StringSplitOptions.RemoveEmptyEntries);
Console.WriteLine(parts.Length);"#
        ),
        &["3"]
    );
}

#[test]
fn split_trim_entries_option() {
    assert_eq!(
        run_csharp(
            r#"var parts=" a , b , c ".Split(',',System.StringSplitOptions.TrimEntries);
Console.WriteLine(parts[1]);"#
        ),
        &["b"]
    );
}

#[test]
fn split_on_string_delimiter() {
    assert_eq!(
        run_csharp(
            r#"var parts="one::two::three".Split("::");
Console.WriteLine(parts.Length); Console.WriteLine(parts[2]);"#
        ),
        &["3", "three"]
    );
}

#[test]
fn join_ienumerable_range() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Join("+",Enumerable.Range(1,4)));"#),
        &["1+2+3+4"]
    );
}
