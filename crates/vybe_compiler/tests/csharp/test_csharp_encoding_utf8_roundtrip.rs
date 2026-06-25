//! `Encoding.UTF8` maps between Unicode strings and byte sequences.
use super::helpers::run_csharp;

#[test]
fn utf8_get_bytes_and_get_string_roundtrip_preserves_text() {
    assert_eq!(
        run_csharp(
            r#"
var encoding = System.Text.Encoding.UTF8;
var bytes = encoding.GetBytes("café");
var text = encoding.GetString(bytes);
Console.WriteLine(text);
"#
        ),
        &["café"]
    );
}
