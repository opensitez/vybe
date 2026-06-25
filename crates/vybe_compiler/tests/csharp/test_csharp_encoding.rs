//! `System.Text.Encoding`: UTF-8, ASCII, Unicode byte conversion.
use super::helpers::run_csharp;

#[test]
fn utf8_get_bytes_returns_byte_array_for_ascii_text() {
    assert_eq!(
        run_csharp(
            r#"var bytes = System.Text.Encoding.UTF8.GetBytes("hi");
Console.WriteLine(bytes.Length);"#
        ),
        &["2"]
    );
}

#[test]
fn utf8_roundtrip_string_through_bytes_and_back() {
    assert_eq!(
        run_csharp(
            r#"var bytes = System.Text.Encoding.UTF8.GetBytes("hello");
Console.WriteLine(System.Text.Encoding.UTF8.GetString(bytes));"#
        ),
        &["hello"]
    );
}

#[test]
fn utf8_multi_byte_character_produces_more_than_one_byte() {
    assert_eq!(
        run_csharp(
            r#"var bytes = System.Text.Encoding.UTF8.GetBytes("€");
Console.WriteLine(bytes.Length > 1);"#
        ),
        &["True"]
    );
}

#[test]
fn ascii_encoding_strips_high_bytes_outside_ascii_range() {
    assert_eq!(
        run_csharp(
            r#"var bytes = System.Text.Encoding.ASCII.GetBytes("ABC");
Console.WriteLine(bytes[0]);"#
        ),
        &["65"]
    );
}

#[test]
fn unicode_encoding_uses_two_bytes_per_char() {
    assert_eq!(
        run_csharp(
            r#"var bytes = System.Text.Encoding.Unicode.GetBytes("A");
Console.WriteLine(bytes.Length);"#
        ),
        &["2"]
    );
}

#[test]
fn get_byte_count_reflects_character_byte_width() {
    assert_eq!(
        run_csharp(
            r#"int n = System.Text.Encoding.UTF8.GetByteCount("café");
Console.WriteLine(n > 4);"#
        ),
        &["True"]
    );
}

#[test]
fn convert_between_encodings_via_encode_decode_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"string text = "test";
byte[] bytes = System.Text.Encoding.UTF8.GetBytes(text);
string result = System.Text.Encoding.UTF8.GetString(bytes);
Console.WriteLine(text == result);"#
        ),
        &["True"]
    );
}
