//! `Convert`, `Uri`, `Path`, and `Encoding` — type bridges and path/URI parsing.
use super::helpers::run_csharp;

#[test]
fn convert_to_boolean_maps_nonzero_integers_to_true() {
    assert_eq!(
        run_csharp(r#"// convert_uri_path
Console.WriteLine(System.Convert.ToBoolean(1));"#),
        &["True"]
    );
}

#[test]
fn convert_to_int32_truncates_double_toward_zero() {
    assert_eq!(
        run_csharp(r#"// convert_uri_path
Console.WriteLine(System.Convert.ToInt32(3.9));"#),
        &["3"]
    );
}

#[test]
fn convert_from_base64_roundtrips_byte_payload() {
    assert_eq!(
        run_csharp(
            r#"
var bytes = System.Convert.FromBase64String("AQID");
Console.WriteLine(bytes.Length);
Console.WriteLine(bytes[2]);
"#
        ),
        &["3", "3"]
    );
}

#[test]
fn uri_absolute_path_excludes_scheme_and_host() {
    assert_eq!(
        run_csharp(
            r#"
var link = new System.Uri("https://example.com/api/v1");
Console.WriteLine(link.AbsolutePath);
"#
        ),
        &["/api/v1"]
    );
}

#[test]
fn uri_combine_resolves_relative_segment_against_base() {
    assert_eq!(
        run_csharp(
            r#"
var baseUri = new System.Uri("https://example.com/a/");
var combined = new System.Uri(baseUri, "b");
Console.WriteLine(combined.AbsolutePath);
"#
        ),
        &["/a/b"]
    );
}

#[test]
fn uri_is_absolute_distinguishes_full_url_from_relative_path() {
    assert_eq!(
        run_csharp(
            r#"
var absolute = new System.Uri("https://example.com");
var relative = new System.Uri("/only-path", System.UriKind.Relative);
Console.WriteLine(absolute.IsAbsoluteUri);
Console.WriteLine(relative.IsAbsoluteUri);
"#
        ),
        &["True", "False"]
    );
}

#[test]
fn path_get_extension_returns_suffix_including_dot() {
    assert_eq!(
        run_csharp(r#"// convert_uri_path
Console.WriteLine(System.IO.Path.GetExtension("archive.tar.gz"));"#),
        &[".gz"]
    );
}

#[test]
fn path_change_extension_replaces_trailing_suffix() {
    assert_eq!(
        run_csharp(r#"// convert_uri_path
Console.WriteLine(System.IO.Path.ChangeExtension("data.txt", ".json"));"#),
        &["data.json"]
    );
}

#[test]
fn encoding_utf8_roundtrips_unicode_text() {
    assert_eq!(
        run_csharp(
            r#"
var bytes = System.Text.Encoding.UTF8.GetBytes("café");
var text = System.Text.Encoding.UTF8.GetString(bytes);
Console.WriteLine(text);
"#
        ),
        &["café"]
    );
}
