//! `primitives/http_session` — session identity and lifecycle.
//!
//! Sessions have no spec surface, so they are a primitive
//! (`documentation/httpserver.md` §4a). One implementation backs PHP
//! `session_*`, Rack's `session` and ASP.NET's `Session`, which is what lets a
//! session survive a PHP → C# → Python call chain in one process.

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_compiler::primitives::{dispatch, http_request_env, http_session};
use vybe_platform_wasi::http as wasi_http;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

/// A VM with a request published the way `vybex --serve` does.
fn vm_with_request(headers: Vec<(String, Vec<u8>)>) -> VM {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let request_id = wasi_http::push_incoming_request(
        "GET",
        Some("/page".to_string()),
        Some("https".to_string()),
        Some("app.test:443".to_string()),
        headers,
        Vec::new(),
    );
    let handle = wasi_http::incoming_request_value(&vm, request_id).expect("request handle");
    vm.globals
        .insert(http_request_env::REQUEST_GLOBAL.to_string(), handle);
    vm
}

/// Emit a sequence of `common:` ops, returning the value left by the last one.
/// Ops ending in `!` have their result dropped.
fn run(vm: &mut VM, ops: &[&str]) -> Value {
    let mut chunks = vec![Chunk::new("<session-test>")];
    for (index, op) in ops.iter().enumerate() {
        let (name, drop) = match op.strip_suffix('!') {
            Some(stripped) => (stripped, true),
            None => (*op, false),
        };
        assert!(
            dispatch::emit_common(name, &mut chunks, 0, 0, 0),
            "common:{name} was not dispatched"
        );
        if drop || index + 1 < ops.len() {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    vm.run(chunks).expect("VM run failed")
}

fn text(value: &Value) -> String {
    match value {
        Value::String(s) => s.to_string(),
        other => panic!("expected a string, got {other:?}"),
    }
}

fn cookie_header(value: &str) -> Vec<(String, Vec<u8>)> {
    vec![("cookie".to_string(), value.as_bytes().to_vec())]
}

#[test]
fn a_new_id_is_32_lowercase_hex_chars() {
    let mut vm = vm_with_request(Vec::new());
    let id = text(&run(&mut vm, &["http_session.new_id"]));
    assert_eq!(id.len(), 32, "got {id:?}");
    assert!(
        id.chars()
            .all(|c| c.is_ascii_hexdigit() && !c.is_uppercase()),
        "got {id:?}"
    );
}

#[test]
fn new_ids_differ() {
    // A session id is a bearer credential; a repeated one is a total
    // authentication bypass, not a cosmetic issue.
    let mut vm = vm_with_request(Vec::new());
    let first = text(&run(&mut vm, &["http_session.new_id"]));
    let second = text(&run(&mut vm, &["http_session.new_id"]));
    assert_ne!(first, second);
}

#[test]
fn the_id_is_adopted_from_the_request_cookie() {
    let mut vm = vm_with_request(cookie_header("SESSIONID=abc123"));
    assert_eq!(text(&run(&mut vm, &["http_session.id"])), "abc123");
}

#[test]
fn the_id_is_stable_within_a_request() {
    // Memoised: PHP asking for the id and C# asking for it must agree, or the
    // two halves of one request write to different sessions.
    let mut vm = vm_with_request(Vec::new());
    let first = text(&run(&mut vm, &["http_session.id"]));
    let second = text(&run(&mut vm, &["http_session.id"]));
    assert_eq!(first, second);
    assert_eq!(first.len(), 32, "with no cookie the id is freshly minted");
}

#[test]
fn an_id_is_minted_when_the_cookie_names_a_different_session() {
    // The cookie is read under the CURRENT session name, not any name.
    let mut vm = vm_with_request(cookie_header("OTHERNAME=abc123"));
    let id = text(&run(&mut vm, &["http_session.id"]));
    assert_ne!(id, "abc123");
    assert_eq!(id.len(), 32);
}

#[test]
fn the_id_is_not_percent_encoded() {
    // Real PHP does not encode session ids, and an encoded id would come back
    // out of session_id() encoded — the value apps put in URLs.
    let mut vm = vm_with_request(Vec::new());
    let id = text(&run(&mut vm, &["http_session.id"]));
    assert!(!id.contains('%'), "got {id:?}");
}

#[test]
fn the_name_falls_back_to_the_languages_spelling() {
    let mut vm = vm_with_request(Vec::new());
    assert_eq!(text(&run(&mut vm, &["http_session.name"])), "SESSIONID");
}

#[test]
fn setting_the_name_changes_which_cookie_is_read() {
    // PHP's session_name('MYAPP') before session_start().
    let mut vm = vm_with_request(cookie_header("MYAPP=fromcookie"));
    let mut chunks = vec![Chunk::new("<session-test>")];
    chunks[0].emit_string_const("MYAPP", 0);
    assert!(dispatch::emit_common(
        "http_session.set_name",
        &mut chunks,
        0,
        1,
        0
    ));
    assert!(dispatch::emit_common(
        "http_session.id",
        &mut chunks,
        0,
        0,
        0
    ));
    chunks[0].emit_op(Op::RETURN, 0);
    assert_eq!(text(&vm.run(chunks).expect("VM run failed")), "fromcookie");
}

#[test]
fn status_starts_as_not_started_and_becomes_active() {
    let mut vm = vm_with_request(Vec::new());
    assert_eq!(
        run(&mut vm, &["http_session.status"]),
        Value::I32(http_session::STATUS_NONE)
    );
    assert_eq!(
        run(&mut vm, &["http_session.start!", "http_session.status"]),
        Value::I32(http_session::STATUS_ACTIVE)
    );
}

#[test]
fn starting_twice_is_a_no_op_that_keeps_the_id() {
    // Frameworks and user code both call session_start(); the second call must
    // not mint a new session and silently drop the first one's data.
    let mut vm = vm_with_request(Vec::new());
    let first = text(&run(&mut vm, &["http_session.start!", "http_session.id"]));
    let second = text(&run(&mut vm, &["http_session.start!", "http_session.id"]));
    assert_eq!(first, second);
}

#[test]
fn the_data_map_is_the_same_object_across_calls() {
    // This is the cross-language guarantee: one map, so a write through one
    // language's alias is visible through another's.
    let mut vm = vm_with_request(Vec::new());
    run(&mut vm, &["http_session.start!"]);

    let mut chunks = vec![Chunk::new("<session-test>")];
    assert!(dispatch::emit_common(
        "http_session.data",
        &mut chunks,
        0,
        0,
        0
    ));
    chunks[0].emit_string_const("user", 0);
    chunks[0].emit_string_const("alice", 0);
    vybe_compiler::primitives::collections::emit_set(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::DROP, 0);
    assert!(dispatch::emit_common(
        "http_session.data",
        &mut chunks,
        0,
        0,
        0
    ));
    chunks[0].emit_string_const("user", 0);
    vybe_compiler::primitives::collections::emit_get(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);

    assert_eq!(
        text(&vm.run(chunks).expect("VM run failed")),
        "alice",
        "a write through one handle must be visible through the next"
    );
}

#[test]
fn regenerate_id_changes_the_id_but_keeps_the_data() {
    // Session fixation defence: the client gets an id it could not have chosen,
    // and stays logged in.
    let mut vm = vm_with_request(Vec::new());
    run(&mut vm, &["http_session.start!"]);

    let mut chunks = vec![Chunk::new("<session-test>")];
    assert!(dispatch::emit_common(
        "http_session.data",
        &mut chunks,
        0,
        0,
        0
    ));
    chunks[0].emit_string_const("user", 0);
    chunks[0].emit_string_const("alice", 0);
    vybe_compiler::primitives::collections::emit_set(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::DROP, 0);
    assert!(dispatch::emit_common(
        "http_session.id",
        &mut chunks,
        0,
        0,
        0
    ));
    chunks[0].emit_op(Op::RETURN, 0);
    let before = text(&vm.run(chunks).expect("VM run failed"));

    let after = text(&run(
        &mut vm,
        &["http_session.regenerate_id!", "http_session.id"],
    ));
    assert_ne!(after, before, "the id must change");
    assert_eq!(after.len(), 32);

    let mut chunks = vec![Chunk::new("<session-test>")];
    assert!(dispatch::emit_common(
        "http_session.data",
        &mut chunks,
        0,
        0,
        0
    ));
    chunks[0].emit_string_const("user", 0);
    vybe_compiler::primitives::collections::emit_get(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    assert_eq!(
        text(&vm.run(chunks).expect("VM run failed")),
        "alice",
        "regenerating the id must not log the user out"
    );
}

#[test]
fn destroy_clears_the_data_and_closes_the_session() {
    let mut vm = vm_with_request(Vec::new());
    run(&mut vm, &["http_session.start!"]);

    let mut chunks = vec![Chunk::new("<session-test>")];
    assert!(dispatch::emit_common(
        "http_session.data",
        &mut chunks,
        0,
        0,
        0
    ));
    chunks[0].emit_string_const("user", 0);
    chunks[0].emit_string_const("alice", 0);
    vybe_compiler::primitives::collections::emit_set(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::DROP, 0);
    assert!(dispatch::emit_common(
        "http_session.destroy",
        &mut chunks,
        0,
        0,
        0
    ));
    chunks[0].emit_op(Op::DROP, 0);
    assert!(dispatch::emit_common(
        "http_session.data",
        &mut chunks,
        0,
        0,
        0
    ));
    chunks[0].emit_op(Op::RETURN, 0);

    let data = vm.run(chunks).expect("VM run failed");
    let Value::Object(object) = &data else {
        panic!("session data is not an object: {data:?}")
    };
    let object = object.lock().unwrap();
    match &object.kind {
        ObjectKind::Map(entries) => assert_eq!(entries.len(), 0, "destroy must empty the session"),
        other => panic!("session data is not a map: {other:?}"),
    }
    drop(object);

    assert_eq!(
        run(&mut vm, &["http_session.status"]),
        Value::I32(http_session::STATUS_NONE)
    );
}

#[test]
fn a_session_works_with_no_request_at_all() {
    // CLI scripts call session_start() too. Nothing here may require a request
    // to exist, or every non-served script dies.
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let started = run(&mut vm, &["http_session.start"]);
    assert_eq!(started, Value::Bool(true));
    assert_eq!(text(&run(&mut vm, &["http_session.id"])).len(), 32);
}
