//! `StringBuilder` – mutable string buffer full API.
use super::helpers::run_csharp;

#[test]
fn append_builds_string_incrementally() {
    assert_eq!(
        run_csharp(
            r#"var sb = new System.Text.StringBuilder();
sb.Append("Hello"); sb.Append(" World");
Console.WriteLine(sb.ToString());"#
        ),
        &["Hello World"]
    );
}

#[test]
fn append_line_adds_content_followed_by_newline() {
    assert_eq!(
        run_csharp(
            r#"var sb = new System.Text.StringBuilder();
sb.AppendLine("line1");
Console.WriteLine(sb.Length > 5);"#
        ),
        &["True"]
    );
}

#[test]
fn insert_places_string_at_given_index() {
    assert_eq!(
        run_csharp(
            r#"var sb = new System.Text.StringBuilder("ac");
sb.Insert(1,"b");
Console.WriteLine(sb.ToString());"#
        ),
        &["abc"]
    );
}

#[test]
fn remove_deletes_character_range_by_start_and_count() {
    assert_eq!(
        run_csharp(
            r#"var sb = new System.Text.StringBuilder("hello");
sb.Remove(1,3);
Console.WriteLine(sb.ToString());"#
        ),
        &["ho"]
    );
}

#[test]
fn replace_substitutes_all_occurrences_in_buffer() {
    assert_eq!(
        run_csharp(
            r#"var sb = new System.Text.StringBuilder("aabbaa");
sb.Replace("aa","X");
Console.WriteLine(sb.ToString());"#
        ),
        &["XbbX"]
    );
}

#[test]
fn clear_resets_length_to_zero() {
    assert_eq!(
        run_csharp(
            r#"var sb = new System.Text.StringBuilder("data");
sb.Clear();
Console.WriteLine(sb.Length);"#
        ),
        &["0"]
    );
}

#[test]
fn length_property_tracks_current_content_size() {
    assert_eq!(
        run_csharp(
            r#"var sb = new System.Text.StringBuilder("abc");
Console.WriteLine(sb.Length);"#
        ),
        &["3"]
    );
}

#[test]
fn indexer_reads_character_at_position() {
    assert_eq!(
        run_csharp(
            r#"var sb = new System.Text.StringBuilder("xyz");
Console.WriteLine(sb[1]);"#
        ),
        &["y"]
    );
}

#[test]
fn append_format_interpolates_value_into_buffer() {
    assert_eq!(
        run_csharp(
            r#"var sb = new System.Text.StringBuilder();
sb.AppendFormat("Value={0}", 42);
Console.WriteLine(sb.ToString());"#
        ),
        &["Value=42"]
    );
}
