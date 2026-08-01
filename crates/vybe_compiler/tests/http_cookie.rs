//! `primitives/http_cookie` — RFC 6265 parsing and serialization.
//!
//! No spec surface provides cookies: `wasi:http` carries headers, and Node's
//! cookie support is userland. So this is a primitive
//! (`documentation/httpserver.md` §4a), emitted once for PHP `$_COOKIE`/
//! `setcookie`, Python `http.cookies`, Rack and ASP.NET alike.

use vybe_compiler::primitives::dispatch;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

fn new_vm() -> VM {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm
}

/// `common:http_cookie.parse` over a `Cookie:` header value.
fn parse(header: &str) -> Value {
    let mut vm = new_vm();
    let mut chunks = vec![Chunk::new("<cookie-test>")];
    let constant = chunks[0].add_constant(Value::String(std::sync::Arc::from(header)));
    chunks[0].emit_op_u16(Op::CONST, constant, 0);
    assert!(dispatch::emit_common(
        "http_cookie.parse",
        &mut chunks,
        0,
        1,
        0
    ));
    chunks[0].emit_op(Op::RETURN, 0);
    vm.run(chunks).expect("VM run failed")
}

/// `common:http_cookie.serialize` with an optional attribute map.
fn serialize(name: &str, value: &str, attrs: &[(&str, Value)]) -> String {
    let mut vm = new_vm();
    let mut chunks = vec![Chunk::new("<cookie-test>")];
    chunks[0].emit_string_const(name, 0);
    chunks[0].emit_string_const(value, 0);

    let argc = if attrs.is_empty() {
        2
    } else {
        // `emit_set` returns Null, not the map, so the map has to be held in a
        // local and re-pushed for each attribute.
        let map = chunks[0].alloc_scratch(1);
        vybe_compiler::primitives::collections::emit_map_new(&mut chunks, 0, 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, map, 0);
        for (key, attr) in attrs {
            chunks[0].emit_op_u16(Op::LOCAL_GET, map, 0);
            chunks[0].emit_string_const(key, 0);
            let constant = chunks[0].add_constant(attr.clone());
            chunks[0].emit_op_u16(Op::CONST, constant, 0);
            vybe_compiler::primitives::collections::emit_set(&mut chunks, 0, 0);
            chunks[0].emit_op(Op::DROP, 0);
        }
        chunks[0].emit_op_u16(Op::LOCAL_GET, map, 0);
        3
    };

    assert!(dispatch::emit_common(
        "http_cookie.serialize",
        &mut chunks,
        0,
        argc,
        0
    ));
    chunks[0].emit_op(Op::RETURN, 0);
    match vm.run(chunks).expect("VM run failed") {
        Value::String(s) => s.to_string(),
        other => panic!("expected a string, got {other:?}"),
    }
}

fn get(map: &Value, name: &str) -> Option<String> {
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
                _ => None,
            }),
        _ => None,
    }
}

fn len(map: &Value) -> usize {
    let Value::Object(object) = map else { return 0 };
    let object = object.lock().unwrap();
    match &object.kind {
        ObjectKind::Map(entries) => entries.len(),
        _ => 0,
    }
}

#[test]
fn parses_a_cookie_pair_list() {
    let jar = parse("a=1; b=2");
    assert_eq!(get(&jar, "a").as_deref(), Some("1"));
    assert_eq!(get(&jar, "b").as_deref(), Some("2"));
    assert_eq!(len(&jar), 2);
}

#[test]
fn a_bare_token_is_not_a_cookie() {
    // RFC 6265 §4.2.1 — a cookie-pair needs an `=`; storing `junk` under an
    // empty name would make `isset($_COOKIE[''])` true.
    let jar = parse("a=1; junk; b=2");
    assert_eq!(len(&jar), 2);
}

#[test]
fn an_absent_cookie_header_is_an_empty_jar() {
    assert_eq!(len(&parse("")), 0);
}

#[test]
fn serializes_a_bare_pair() {
    assert_eq!(serialize("name", "value", &[]), "name=value");
}

#[test]
fn the_value_is_written_verbatim() {
    // Encoding is the calling language's choice — PHP's setcookie() encodes and
    // setrawcookie() does not — so the primitive must not encode.
    assert_eq!(serialize("k", "a b+c%2D", &[]), "k=a b+c%2D");
}

#[test]
fn serializes_path_domain_and_samesite() {
    let out = serialize(
        "k",
        "v",
        &[
            ("path", Value::String("/app".into())),
            ("domain", Value::String("example.test".into())),
            ("samesite", Value::String("Lax".into())),
        ],
    );
    assert_eq!(out, "k=v; Path=/app; Domain=example.test; SameSite=Lax");
}

#[test]
fn flags_are_valueless_and_only_appear_when_true() {
    assert_eq!(
        serialize(
            "k",
            "v",
            &[
                ("secure", Value::Bool(true)),
                ("httponly", Value::Bool(true))
            ],
        ),
        "k=v; Secure; HttpOnly"
    );
    assert_eq!(
        serialize(
            "k",
            "v",
            &[
                ("secure", Value::Bool(false)),
                ("httponly", Value::Bool(false))
            ],
        ),
        "k=v",
        "a false flag must be omitted, not written as `; Secure`"
    );
}

#[test]
fn expires_becomes_an_http_date() {
    // Callers pass unix seconds; `Expires` is an IMF-fixdate (RFC 9110 §5.6.7).
    // Emitting the raw timestamp would make the cookie never expire.
    let out = serialize("k", "v", &[("expires", Value::F64(1_754_006_400.0))]);
    assert!(
        out.starts_with("k=v; Expires="),
        "got {out:?}"
    );
    assert!(out.ends_with(" GMT"), "got {out:?}");
    assert!(
        !out.contains("1754006400"),
        "the unix timestamp must not survive into the header: {out:?}"
    );
}

#[test]
fn a_zero_expiry_is_a_session_cookie() {
    // PHP uses `expires = 0` for "until the browser closes" — that must NOT
    // become `Expires=Thu, 01 Jan 1970`, which would delete the cookie instead.
    assert_eq!(serialize("k", "v", &[("expires", Value::F64(0.0))]), "k=v");
}

#[test]
fn a_round_trip_survives_parsing() {
    let serialized = serialize("sid", "abc123", &[("path", Value::String("/".into()))]);
    let pair = serialized.split(';').next().expect("a cookie-pair");
    let jar = parse(pair);
    assert_eq!(get(&jar, "sid").as_deref(), Some("abc123"));
}
