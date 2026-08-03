//! `primitives/http_form` — request body parsing, shared across languages.
//!
//! Backs PHP `$_POST`, Python `cgi.FieldStorage`, Rack's form parsing and
//! ASP.NET's form collection. One implementation, per
//! `documentation/httpserver.md` §4a.

use vybe_compiler::primitives::dispatch;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

/// Run `common:http_form.parse_urlencoded` over `body`.
fn parse(body: &str) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());

    let mut chunks = vec![Chunk::new("<http-form-test>")];
    let constant = chunks[0].add_constant(Value::String(std::sync::Arc::from(body)));
    chunks[0].emit_op_u16(Op::CONST, constant, 0);
    assert!(dispatch::emit_common(
        "http_form.parse_urlencoded",
        &mut chunks,
        0,
        1,
        0
    ));
    chunks[0].emit_op(Op::RETURN, 0);
    vm.run(chunks).expect("VM run failed")
}

fn key(map: &Value, name: &str) -> Option<String> {
    let Value::Object(object) = map else {
        return None;
    };
    let object = object.lock().unwrap();
    match &object.kind {
        ObjectKind::Map(entries) => entries
            .iter()
            .find(|(k, _)| matches!(k, Value::String(s) if s.as_ref() == name))
            .and_then(|(_, v)| match v {
                Value::String(s) => Some(s.to_string()),
                _ => None }),
        _ => None }
}

fn len(map: &Value) -> usize {
    let Value::Object(object) = map else { return 0 };
    let object = object.lock().unwrap();
    match &object.kind {
        ObjectKind::Map(entries) => entries.len(),
        _ => 0 }
}

#[test]
fn parses_simple_pairs() {
    let form = parse("a=1&b=2");
    assert_eq!(key(&form, "a").as_deref(), Some("1"));
    assert_eq!(key(&form, "b").as_deref(), Some("2"));
    assert_eq!(len(&form), 2);
}

#[test]
fn plus_decodes_to_space() {
    // RFC 1866 §8.2.1 — `+` is a space in urlencoded, unlike in a URI path.
    let form = parse("name=John+Smith");
    assert_eq!(key(&form, "name").as_deref(), Some("John Smith"));
}

#[test]
fn percent_escapes_are_decoded() {
    let form = parse("q=caf%C3%A9&op=%3D");
    assert_eq!(key(&form, "q").as_deref(), Some("café"));
    assert_eq!(key(&form, "op").as_deref(), Some("="));
}

#[test]
fn keys_are_decoded_too() {
    let form = parse("user%20name=x");
    assert_eq!(key(&form, "user name").as_deref(), Some("x"));
}

#[test]
fn an_empty_body_is_an_empty_map() {
    // No body must not be an error — every caller treats $_POST/environ as a
    // map that is simply empty for GET.
    assert_eq!(len(&parse("")), 0);
}

#[test]
fn segments_without_an_equals_are_skipped() {
    // `a=1&junk&b=2` — a bare token is not a field, and must not land under an
    // empty key.
    let form = parse("a=1&junk&b=2");
    assert_eq!(key(&form, "a").as_deref(), Some("1"));
    assert_eq!(key(&form, "b").as_deref(), Some("2"));
    assert_eq!(len(&form), 2, "bare token should not become a field");
}

#[test]
fn an_empty_value_is_kept() {
    // `a=` is a present field with an empty value — distinct from absent.
    let form = parse("a=&b=2");
    assert_eq!(key(&form, "a").as_deref(), Some(""));
    assert_eq!(len(&form), 2);
}

/// Run `common:http_form.parse_multipart` over `body` with `content_type`.
fn parse_multipart(body: &str, content_type: &str) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());

    let mut chunks = vec![Chunk::new("<http-form-test>")];
    let body_const = chunks[0].add_constant(Value::String(std::sync::Arc::from(body)));
    chunks[0].emit_op_u16(Op::CONST, body_const, 0);
    let ct_const = chunks[0].add_constant(Value::String(std::sync::Arc::from(content_type)));
    chunks[0].emit_op_u16(Op::CONST, ct_const, 0);
    assert!(dispatch::emit_common(
        "http_form.parse_multipart",
        &mut chunks,
        0,
        2,
        0
    ));
    chunks[0].emit_op(Op::RETURN, 0);
    vm.run(chunks).expect("VM run failed")
}

/// The `fields` or `files` sub-map of a multipart result.
fn sub(map: &Value, which: &str) -> Value {
    let Value::Object(object) = map else {
        panic!("multipart result is not an object")
    };
    let object = object.lock().unwrap();
    match &object.kind {
        ObjectKind::Map(entries) => entries
            .iter()
            .find(|(k, _)| matches!(k, Value::String(s) if s.as_ref() == which))
            .map(|(_, v)| v.clone())
            .unwrap_or_else(|| panic!("no `{which}` in multipart result")),
        _ => panic!("multipart result is not a map") }
}

/// Wrap parts in boundary delimiters. Parts are passed already CRLF-correct —
/// no line-ending rewriting here, since a body may carry binary content whose
/// bare `0x0A` bytes must not become CRLF.
/// One field of one upload: `files[name][field]`.
///
/// An upload is a MAP, not a bare string — the client filename, the declared
/// media type and the octet count cannot be recovered from the bytes, and
/// every language surfaces all three (`$_FILES[k]['name']`, Rack `:filename`).
fn upload(form: &Value, name: &str, field: &str) -> Option<String> {
    let files = sub(form, "files");
    let Value::Object(object) = &files else {
        return None;
    };
    let entry = {
        let guard = object.lock().unwrap();
        match &guard.kind {
            ObjectKind::Map(entries) => entries
                .iter()
                .find(|(k, _)| matches!(k, Value::String(s) if s.as_ref() == name))
                .map(|(_, v)| v.clone()),
            _ => None }
    }?;
    key(&entry, field)
}

fn multipart(parts: &[&str]) -> String {
    let mut body = String::new();
    for part in parts {
        body.push_str("--BOUNDARY\r\n");
        body.push_str(part);
        body.push_str("\r\n");
    }
    body.push_str("--BOUNDARY--\r\n");
    body
}

const CT: &str = "multipart/form-data; boundary=BOUNDARY";

#[test]
fn multipart_reads_a_plain_field() {
    let form = parse_multipart(
        &multipart(&["Content-Disposition: form-data; name=\"a\"\r\n\r\n1"]),
        CT,
    );
    assert_eq!(key(&sub(&form, "fields"), "a").as_deref(), Some("1"));
    assert_eq!(len(&sub(&form, "files")), 0);
}

#[test]
fn multipart_separates_files_from_fields() {
    // A part with `filename` is an upload; without it, a field. Every language
    // surfaces the two separately ($_POST vs $_FILES), so the split is here.
    let form = parse_multipart(
        &multipart(&[
            "Content-Disposition: form-data; name=\"note\"\r\n\r\nhello",
            "Content-Disposition: form-data; name=\"doc\"; filename=\"a.txt\"\r\nContent-Type: text/plain\r\n\r\nFILE BODY",
        ]),
        CT,
    );
    assert_eq!(key(&sub(&form, "fields"), "note").as_deref(), Some("hello"));
    assert_eq!(
        len(&sub(&form, "fields")),
        1,
        "the upload must not be a field"
    );
    assert_eq!(
        upload(&form, "doc", "content").as_deref(),
        Some("FILE BODY")
    );
    assert_eq!(upload(&form, "doc", "filename").as_deref(), Some("a.txt"));
    assert_eq!(upload(&form, "doc", "type").as_deref(), Some("text/plain"));
}

#[test]
fn multipart_keeps_content_verbatim() {
    // No decoding in multipart — `+`, `%` and `=` are literal, unlike
    // urlencoded. Getting this wrong silently corrupts every upload.
    let form = parse_multipart(
        &multipart(&["Content-Disposition: form-data; name=\"raw\"\r\n\r\na+b %41 c=d"]),
        CT,
    );
    assert_eq!(
        key(&sub(&form, "fields"), "raw").as_deref(),
        Some("a+b %41 c=d")
    );
}

#[test]
fn multipart_content_may_contain_blank_lines() {
    // Only the FIRST blank line ends the headers; later ones are content.
    let form = parse_multipart(
        &multipart(&["Content-Disposition: form-data; name=\"t\"\r\n\r\nline1\r\n\r\nline3"]),
        CT,
    );
    assert_eq!(
        key(&sub(&form, "fields"), "t").as_deref(),
        Some("line1\r\n\r\nline3")
    );
}

#[test]
fn multipart_keeps_an_empty_upload() {
    // A zero-byte upload is still an upload — PHP reports it with an error
    // code, it does not vanish.
    let form = parse_multipart(
        &multipart(&["Content-Disposition: form-data; name=\"f\"; filename=\"empty.txt\"\r\n\r\n"]),
        CT,
    );
    assert_eq!(upload(&form, "f", "content").as_deref(), Some(""));
    assert_eq!(len(&sub(&form, "files")), 1);
}

#[test]
fn multipart_ignores_preamble_and_epilogue() {
    // Splitting on the delimiter yields a leading segment before the first
    // boundary and a trailing `--`; neither is a part.
    let mut body = String::from("ignored preamble\r\n");
    body.push_str(&multipart(&[
        "Content-Disposition: form-data; name=\"a\"\r\n\r\n1",
    ]));
    let form = parse_multipart(&body, CT);
    assert_eq!(len(&sub(&form, "fields")), 1);
    assert_eq!(key(&sub(&form, "fields"), "a").as_deref(), Some("1"));
}

#[test]
fn multipart_accepts_a_quoted_boundary() {
    // RFC 2045 §5.1 — boundary is a quoted-string when it contains specials.
    let form = parse_multipart(
        &multipart(&["Content-Disposition: form-data; name=\"a\"\r\n\r\n1"]),
        "multipart/form-data; boundary=\"BOUNDARY\"",
    );
    assert_eq!(key(&sub(&form, "fields"), "a").as_deref(), Some("1"));
}

#[test]
fn multipart_reads_a_boundary_that_is_not_last() {
    // `boundary=` is not always the final parameter.
    let form = parse_multipart(
        &multipart(&["Content-Disposition: form-data; name=\"a\"\r\n\r\n1"]),
        "multipart/form-data; boundary=BOUNDARY; charset=utf-8",
    );
    assert_eq!(key(&sub(&form, "fields"), "a").as_deref(), Some("1"));
}

#[test]
fn multipart_skips_a_part_with_no_name() {
    let form = parse_multipart(&multipart(&["Content-Type: text/plain\r\n\r\norphan"]), CT);
    assert_eq!(len(&sub(&form, "fields")), 0);
    assert_eq!(len(&sub(&form, "files")), 0);
}

#[test]
fn multipart_binary_content_round_trips() {
    // The body is handled as LATIN-1 (one char per byte) so binary uploads
    // survive; a UTF-8 decode here would mangle every image.
    let raw: String = (0u8..=255).map(|b| b as char).collect();
    let form = parse_multipart(
        &multipart(&[&format!(
            "Content-Disposition: form-data; name=\"b\"; filename=\"x.bin\"\r\n\r\n{raw}"
        )]),
        CT,
    );
    let got = upload(&form, "b", "content").expect("upload missing");
    assert_eq!(got.chars().count(), 256);
    assert_eq!(got.chars().next(), Some('\u{0}'));
    assert_eq!(got.chars().last(), Some('\u{ff}'));
}

#[test]
fn multipart_accepts_filename_before_name() {
    // RFC 7578 fixes no parameter order. A bare substring search finds `name=`
    // inside `filename=`, filing the upload under the file's name.
    let form = parse_multipart(
        &multipart(&["Content-Disposition: form-data; filename=\"x.txt\"; name=\"f\"\r\n\r\nBODY"]),
        CT,
    );
    assert_eq!(upload(&form, "f", "content").as_deref(), Some("BODY"));
    assert_eq!(
        len(&sub(&form, "fields")),
        0,
        "must not be filed as a field"
    );
}

#[test]
fn multipart_accepts_a_parameter_with_no_space_after_the_semicolon() {
    let form = parse_multipart(
        &multipart(&["Content-Disposition: form-data;name=\"a\"\r\n\r\n1"]),
        CT,
    );
    assert_eq!(key(&sub(&form, "fields"), "a").as_deref(), Some("1"));
}
