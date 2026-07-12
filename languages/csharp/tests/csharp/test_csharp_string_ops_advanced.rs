//! String operations not covered elsewhere: spans, split edge cases, format, interpolation.
use super::helpers::run_csharp;

#[test]
fn split_with_multiple_delimiters() {
    assert_eq!(
        run_csharp(
            r#"var parts="a,b;c".Split(new char[]{',',';'});
Console.WriteLine(parts.Length); Console.WriteLine(parts[2]);"#
        ),
        &["3", "c"]
    );
}

#[test]
fn split_with_max_count_limits_resulting_segments() {
    assert_eq!(
        run_csharp(
            r#"var parts="a:b:c:d".Split(':',2);
Console.WriteLine(parts.Length); Console.WriteLine(parts[1]);"#
        ),
        &["2", "b:c:d"]
    );
}

#[test]
fn trim_start_removes_only_leading_whitespace() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("  hi  ".TrimStart());"#),
        &["hi  "]
    );
}

#[test]
fn trim_end_removes_only_trailing_whitespace() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("  hi  ".TrimEnd());"#),
        &["  hi"]
    );
}

#[test]
fn substring_single_arg_returns_suffix_from_index() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("hello world".Substring(6));"#),
        &["world"]
    );
}

#[test]
fn string_repeat_via_new_string_ctor() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new string('*', 5));"#),
        &["*****"]
    );
}

#[test]
fn string_to_char_array_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"char[] chars="abc".ToCharArray();
Console.WriteLine(new string(chars));"#
        ),
        &["abc"]
    );
}

#[test]
fn string_insert_inserts_at_position() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("helo".Insert(3,"l"));"#),
        &["hello"]
    );
}

#[test]
fn string_remove_deletes_range() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("hello world".Remove(5,6));"#),
        &["hello"]
    );
}

#[test]
fn string_format_with_named_composite_via_positional() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Format("{0:000}", 7));"#),
        &["007"]
    );
}

#[test]
fn interpolated_string_with_format_specifier() {
    assert_eq!(
        run_csharp(r#"double pi=3.14159; Console.WriteLine($"{pi:F2}");"#),
        &["3.14"]
    );
}

#[test]
fn interpolated_string_with_ternary_expression() {
    assert_eq!(
        run_csharp(r#"int n=5; Console.WriteLine($"{(n>3?"big":"small")}");"#),
        &["big"]
    );
}
