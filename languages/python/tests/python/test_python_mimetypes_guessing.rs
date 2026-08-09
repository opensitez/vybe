use super::helpers::run_python;

// mimetypes — guess_type, guess_extension, add_type, MimeTypes class

#[test]
fn test_mimetypes_guess_html_type() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("file.html")
print(t)
print(enc)
"#,
    );
    assert_eq!(out, vec!["text/html", "None"]);
}

#[test]
fn test_mimetypes_guess_jpeg_type() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("photo.jpg")
print(t)
"#,
    );
    assert_eq!(out, vec!["image/jpeg"]);
}

#[test]
fn test_mimetypes_guess_gzip_encoding() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("archive.tar.gz")
print(enc)
"#,
    );
    assert_eq!(out, vec!["gzip"]);
}

#[test]
fn test_mimetypes_guess_json_type() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("data.json")
print(t)
"#,
    );
    assert_eq!(out, vec!["application/json"]);
}

#[test]
fn test_mimetypes_guess_css_type() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("style.css")
print(t)
"#,
    );
    assert_eq!(out, vec!["text/css"]);
}

#[test]
fn test_mimetypes_guess_png_type() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("image.png")
print(t)
"#,
    );
    assert_eq!(out, vec!["image/png"]);
}

#[test]
fn test_mimetypes_guess_unknown_type_returns_none() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("file.unknownxyz123")
print(t)
"#,
    );
    assert_eq!(out, vec!["None"]);
}

#[test]
fn test_mimetypes_guess_extension_text_html() {
    let out = run_python(
        r#"
import mimetypes
ext = mimetypes.guess_extension("text/html")
print(ext in [".html", ".htm"])
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_mimetypes_guess_all_extensions_text_html() {
    let out = run_python(
        r#"
import mimetypes
exts = mimetypes.guess_all_extensions("text/html")
print(".html" in exts or ".htm" in exts)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_mimetypes_add_type_custom() {
    let out = run_python(
        r#"
import mimetypes
mimetypes.add_type("application/x-custom-vybe", ".vyb")
t, enc = mimetypes.guess_type("test.vyb")
print(t)
"#,
    );
    assert_eq!(out, vec!["application/x-custom-vybe"]);
}

#[test]
fn test_mimetypes_types_map_contains_html() {
    let out = run_python(
        r#"
import mimetypes
mimetypes.init()
print(".html" in mimetypes.types_map or ".htm" in mimetypes.types_map)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_mimetypes_encodings_map_gzip() {
    let out = run_python(
        r#"
import mimetypes
print(".gz" in mimetypes.encodings_map)
print(mimetypes.encodings_map[".gz"])
"#,
    );
    assert_eq!(out, vec!["True", "gzip"]);
}

#[test]
fn test_mimetypes_suffix_map_tgz() {
    let out = run_python(
        r#"
import mimetypes
print(".tgz" in mimetypes.suffix_map)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_mimetypes_mimetype_class_read() {
    let out = run_python(
        r#"
import mimetypes
mt = mimetypes.MimeTypes()
t, enc = mt.guess_type("index.html")
print(t)
"#,
    );
    assert_eq!(out, vec!["text/html"]);
}

#[test]
fn test_mimetypes_guess_javascript_type() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("script.js")
print(t in ["text/javascript", "application/javascript"])
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_mimetypes_guess_xml_type() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("data.xml")
print(t in ["text/xml", "application/xml"])
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_mimetypes_guess_pdf_type() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("doc.pdf")
print(t)
"#,
    );
    assert_eq!(out, vec!["application/pdf"]);
}

#[test]
fn test_mimetypes_guess_type_with_url() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("https://example.com/file.png")
print(t)
"#,
    );
    assert_eq!(out, vec!["image/png"]);
}

#[test]
fn test_mimetypes_guess_extension_returns_none_for_unknown() {
    let out = run_python(
        r#"
import mimetypes
ext = mimetypes.guess_extension("application/x-totally-unknown-type-xyzabc")
print(ext)
"#,
    );
    assert_eq!(out, vec!["None"]);
}

#[test]
fn test_mimetypes_guess_zip_type() {
    let out = run_python(
        r#"
import mimetypes
t, enc = mimetypes.guess_type("archive.zip")
print(t)
"#,
    );
    assert_eq!(out, vec!["application/zip"]);
}
