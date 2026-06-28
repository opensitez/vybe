//! Core `string` instance and static methods.
use super::helpers::run_csharp;

#[test]
fn split_divides_on_single_char_delimiter() {
    assert_eq!(
        run_csharp(r#"var p = "a,b,c".Split(','); Console.WriteLine(p[1]);"#),
        &["b"]
    );
}

#[test]
fn split_with_remove_empty_entries_drops_consecutive_delimiters() {
    assert_eq!(
        run_csharp(
            r#"var p = "a,,b".Split(new[]{','}, System.StringSplitOptions.RemoveEmptyEntries);
Console.WriteLine(p.Length);"#
        ),
        &["2"]
    );
}

#[test]
fn join_concatenates_sequence_with_separator() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Join("-", new[]{"a","b","c"}));"#),
        &["a-b-c"]
    );
}

#[test]
fn trim_removes_leading_and_trailing_whitespace() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("  hello  ".Trim());"#),
        &["hello"]
    );
}

#[test]
fn pad_left_right_align_string_to_minimum_width() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("hi".PadLeft(5));"#),
        &["   hi"]
    );
}

#[test]
fn pad_right_left_aligns_to_minimum_width() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("hi".PadRight(5) + "|");"#),
        &["hi   |"]
    );
}

#[test]
fn contains_returns_true_for_present_substring() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("foobar".Contains("oba"));"#),
        &["True"]
    );
}

#[test]
fn starts_with_checks_prefix_match() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("prefix_body".StartsWith("prefix"));"#),
        &["True"]
    );
}

#[test]
fn ends_with_checks_suffix_match() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("body_suffix".EndsWith("suffix"));"#),
        &["True"]
    );
}

#[test]
fn index_of_returns_first_occurrence_position() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("abcabc".IndexOf('b'));"#),
        &["1"]
    );
}

#[test]
fn last_index_of_returns_final_occurrence_position() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("abcabc".LastIndexOf('b'));"#),
        &["4"]
    );
}

#[test]
fn replace_substitutes_all_occurrences_of_substring() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("aabbaa".Replace("aa","X"));"#),
        &["XbbX"]
    );
}

#[test]
fn substring_extracts_region_by_start_and_length() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("hello world".Substring(6, 5));"#),
        &["world"]
    );
}

#[test]
fn to_upper_lower_changes_case_of_all_letters() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine("Hello".ToUpper()); Console.WriteLine("Hello".ToLower());"#
        ),
        &["HELLO", "hello"]
    );
}

#[test]
fn is_null_or_whitespace_returns_true_for_spaces_only() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.IsNullOrWhiteSpace("   "));"#),
        &["True"]
    );
}

#[test]
fn concat_static_joins_array_of_strings_without_separator() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Concat("a","b","c"));"#),
        &["abc"]
    );
}

#[test]
fn compare_ordinal_distinguishes_case_when_flag_false() {
    assert_eq!(
        run_csharp(
            r#"int r = string.Compare("A","a",System.StringComparison.Ordinal);
Console.WriteLine(r < 0 || r > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn string_chars_indexer_reads_individual_character() {
    assert_eq!(run_csharp(r#"Console.WriteLine("hello"[1]);"#), &["e"]);
}
