//! Verbatim `@""` strings: backslash literals and doubled-quote escapes.
use super::helpers::run_csharp;

#[test]
fn verbatim_string_preserves_backslash_characters() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(@"C:\temp\file");"#),
        &[r"C:\temp\file"]
    );
}

#[test]
fn verbatim_string_doubled_quote_embeds_single_quote() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(@"say ""hi""");"#),
        &[r#"say "hi""#]
    );
}

#[test]
fn verbatim_string_spans_multiple_lines_when_source_contains_newlines() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine(@"line1
line2");"#
        ),
        &["line1\nline2"]
    );
}

#[test]
fn regular_string_escape_newline_differs_from_verbatim_multiline() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("a\nb");"#),
        &["a\nb"]
    );
}

#[test]
fn verbatim_empty_string_has_zero_length() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(@"".Length);"#),
        &["0"]
    );
}

#[test]
fn verbatim_string_concatenated_with_regular_string() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(@"dir\" + "name");"#),
        &[r"dir\name"]
    );
}

#[test]
fn verbatim_string_indexer_reads_code_units_same_as_normal_string() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(@"abc"[1]);"#),
        &["b"]
    );
}

#[test]
fn verbatim_string_length_counts_all_characters_including_escapes_as_literals() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(@"\\".Length);"#),
        &["2"]
    );
}

#[test]
fn verbatim_string_equality_compares_literal_content() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(@"x" == "x");"#),
        &["True"]
    );
}

#[test]
fn verbatim_string_starts_with_prefix_when_using_starts_with_method() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(@"C:\data".StartsWith(@"C:\"));"#),
        &["True"]
    );
}
