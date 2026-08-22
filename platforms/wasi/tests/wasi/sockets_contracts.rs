//! Behavioural contracts for `wasi:sockets@0.3.1`.
//!
//! This file used to pin the 0.2 surface: `start-bind`/`finish-bind`,
//! `start-listen`/`finish-listen`, `start-connect`/`finish-connect`, `accept`,
//! `shutdown`, `subscribe`, `udp.stream` and `instance-network`. Every one of
//! those is DELETED in 0.3.1, so every one of those tests was pinning a name a
//! conforming guest cannot import — the same "assert the old spelling still
//! works" lock that kept four other packages from moving.
//!
//! The contracts themselves were worth keeping, so they are restated against
//! the surviving functions. Two behavioural changes are asserted here rather
//! than assumed:
//!
//!   * a failure is an ERROR RECORD (`__wasi_error: <code>`), not `Value::Null`
//!     — 0.2's helpers returned a bare null, 0.3's return `result<_, error-code>`;
//!   * `resolve-addresses` answers `list<ip-address>` DIRECTLY. 0.2 answered a
//!     `resolve-address-stream` resource that had to be drained.
//!
//! Tests for functions with no 0.3.1 counterpart at all (`subscribe`, which
//! went with `wasi:io`'s pollables; `shutdown` and `accept`, which the
//! Component Model replaced with resource-drop and a `stream<tcp-socket>`) are
//! not restated, because there is nothing left to assert about them.

use std::net::TcpListener;
use std::sync::Arc;

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call_import(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-sockets-contracts-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(text) => chunk.emit_string_const(&text, 0),
            Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
            other => {
                let global_name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                vm.set_global_owned(global_name.clone(), other);
                let ci = chunk.intern_string_constant(&global_name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn len_of(value: &Value) -> usize {
    let Value::Object(object) = value else {
        return 0;
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(elements) = &object.kind else {
        return 0;
    };
    elements.len()
}

/// Every variant of `variant error-code`, verbatim from
/// `proposals/WASI/proposals/sockets/wit/types.wit:18`.
///
/// ⛔ `would-block` is NOT here. 0.2 had it; 0.3.1 DELETED it, because
/// `connect`/`send`/`receive` became `async func` and "not ready yet" is the
/// future not having resolved rather than an error. A host that answers
/// `would-block` has invented a code, which is the same offence as inventing a
/// verb — and far easier to miss, because nothing about the call looks wrong.
const ERROR_CODES: &[&str] = &[
    "access-denied",
    "not-supported",
    "invalid-argument",
    "out-of-memory",
    "timeout",
    "invalid-state",
    "address-not-bindable",
    "address-in-use",
    "remote-unreachable",
    "connection-refused",
    "connection-broken",
    "connection-reset",
    "connection-aborted",
    "datagram-too-large",
    "other",
];

/// The `error-code` of a `result<_, error-code>` that went the error way, or
/// `None` when the call succeeded.
///
/// This is the assertion 0.2 could not make: its helpers answered a bare
/// `Value::Null` for every kind of failure, so a test could confirm that
/// something went wrong but never WHAT.
///
/// It also GATES the code against the WIT, so every assertion in this file
/// doubles as a check that the host did not invent a variant — the failure
/// mode that a per-test `assert_eq!(code, "...")` cannot see, because a test
/// pinning an invented code passes happily forever.
fn err_code(value: &Value) -> Option<String> {
    let Value::Object(object) = value else {
        return None;
    };
    let object = object.lock().unwrap();
    match object.properties.get("__wasi_error") {
        Some(Value::String(code)) => {
            let code = code.to_string();
            assert!(
                ERROR_CODES.contains(&code.as_str()),
                "`{code}` is not a variant of wasi:sockets/types error-code; \
                 declared variants are: {}",
                ERROR_CODES.join(", ")
            );
            Some(code)
        }
        _ => None,
    }
}

fn is_ok(value: &Value) -> bool {
    err_code(value).is_none()
}

fn free_port() -> u16 {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind ephemeral port");
    listener.local_addr().expect("local addr").port()
}

fn tcp_socket() -> Value {
    let socket = call_import(
        "wasi:sockets/types",
        "[static]tcp-socket.create",
        vec![s("ipv4")],
    );
    assert!(is_ok(&socket), "create must succeed: {:?}", err_code(&socket));
    socket
}

fn udp_socket() -> Value {
    let socket = call_import(
        "wasi:sockets/types",
        "[static]udp-socket.create",
        vec![s("ipv4")],
    );
    assert!(is_ok(&socket), "create must succeed: {:?}", err_code(&socket));
    socket
}

/// `create` + `bind` + `listen` — three calls where 0.2 needed five, because
/// the `start-*`/`finish-*` pairs collapsed and the `network` handle is gone.
fn listen_on_localhost(port: u16) -> Value {
    let listener = tcp_socket();
    let bound = call_import(
        "wasi:sockets/types",
        "[method]tcp-socket.bind",
        vec![listener.clone(), s(&format!("127.0.0.1:{}", port))],
    );
    assert!(is_ok(&bound), "bind failed: {:?}", err_code(&bound));
    let stream = call_import(
        "wasi:sockets/types",
        "[method]tcp-socket.listen",
        vec![listener.clone()],
    );
    assert!(is_ok(&stream), "listen failed: {:?}", err_code(&stream));
    listener
}

// ── tcp-socket ──────────────────────────────────────────────────────────────

#[test]
fn tcp_get_is_listening_is_false_before_listen() {
    assert_eq!(
        call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.get-is-listening",
            vec![tcp_socket()]
        ),
        Value::Bool(false)
    );
}

#[test]
fn tcp_get_is_listening_is_true_after_listen() {
    let listener = listen_on_localhost(free_port());
    assert_eq!(
        call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.get-is-listening",
            vec![listener]
        ),
        Value::Bool(true)
    );
}

#[test]
fn tcp_get_local_address_on_invalid_handle_is_invalid_state() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.get-local-address",
            vec![Value::Null]
        )),
        Some("invalid-state".to_string())
    );
}

#[test]
fn tcp_get_remote_address_on_invalid_handle_is_invalid_state() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.get-remote-address",
            vec![Value::Null]
        )),
        Some("invalid-state".to_string())
    );
}

#[test]
fn tcp_bind_requires_socket_handle() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.bind",
            vec![Value::Null, s("127.0.0.1:0")]
        )),
        Some("invalid-state".to_string())
    );
}

#[test]
fn tcp_bind_with_missing_address_is_invalid_argument() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.bind",
            vec![tcp_socket()]
        )),
        Some("invalid-argument".to_string())
    );
}

/// "If the port is zero the socket is bound to a random free port" — so the
/// address read back must be the BOUND one, never the requested one.
#[test]
fn tcp_bind_to_port_zero_reports_the_port_actually_bound() {
    let socket = tcp_socket();
    let bound = call_import(
        "wasi:sockets/types",
        "[method]tcp-socket.bind",
        vec![socket.clone(), s("127.0.0.1:0")],
    );
    assert!(is_ok(&bound), "bind failed: {:?}", err_code(&bound));
    let local = call_import(
        "wasi:sockets/types",
        "[method]tcp-socket.get-local-address",
        vec![socket],
    );
    assert!(is_ok(&local), "get-local-address failed");

    // `ip-socket-address` is a RECORD (`family` / `port` / `address`), so the
    // port has to be read out of it. Stringifying and matching on ":0" would
    // pass no matter what the kernel assigned — the Display form is the object
    // rendering, not "host:port", so the suffix never matches either way.
    let Value::Object(object) = &local else {
        panic!("get-local-address must answer an ip-socket-address record");
    };
    let port = object
        .lock()
        .unwrap()
        .properties
        .get("port")
        .map(|value| value.as_f64())
        .expect("ip-socket-address carries a `port` field");
    assert!(
        port > 0.0,
        "port 0 must resolve to the port actually bound, got {port}"
    );
}

#[test]
fn tcp_listen_requires_socket_handle() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.listen",
            vec![Value::Null]
        )),
        Some("invalid-state".to_string())
    );
}

#[test]
fn tcp_connect_requires_socket_handle() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.connect",
            vec![Value::Null, s("127.0.0.1:1")]
        )),
        Some("invalid-state".to_string())
    );
}

#[test]
fn tcp_connect_with_missing_address_is_invalid_argument() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.connect",
            vec![tcp_socket()]
        )),
        Some("invalid-argument".to_string())
    );
}

#[test]
fn tcp_set_listen_backlog_requires_socket_handle() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.set-listen-backlog-size",
            vec![Value::Null, Value::I64(32)]
        )),
        Some("invalid-state".to_string())
    );
}

#[test]
fn tcp_set_listen_backlog_rejects_zero() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]tcp-socket.set-listen-backlog-size",
            vec![tcp_socket(), Value::I64(0)]
        )),
        Some("invalid-argument".to_string())
    );
}

/// `connect` reports the addresses; the byte streams come from `receive`, NOT
/// from a `finish-connect` return value — 0.3.1 has no `finish-connect`.
#[test]
fn tcp_connect_populates_addresses_and_receive_answers_a_stream_pair() {
    let port = free_port();
    let _listener = listen_on_localhost(port);
    let client = tcp_socket();
    let connected = call_import(
        "wasi:sockets/types",
        "[method]tcp-socket.connect",
        vec![client.clone(), s(&format!("127.0.0.1:{}", port))],
    );
    assert!(
        is_ok(&connected),
        "connect failed: {:?}",
        err_code(&connected)
    );

    let remote = call_import(
        "wasi:sockets/types",
        "[method]tcp-socket.get-remote-address",
        vec![client.clone()],
    );
    assert!(is_ok(&remote), "remote address must be known after connect");

    // `receive: func() -> tuple<stream<u8>, future<result<_, error-code>>>`
    let received = call_import(
        "wasi:sockets/types",
        "[method]tcp-socket.receive",
        vec![client],
    );
    assert_eq!(
        len_of(&received),
        2,
        "receive answers a (stream, future) tuple"
    );
}

// ── udp-socket ──────────────────────────────────────────────────────────────

#[test]
fn udp_get_local_address_on_invalid_handle_is_invalid_state() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]udp-socket.get-local-address",
            vec![Value::Null]
        )),
        Some("invalid-state".to_string())
    );
}

#[test]
fn udp_get_remote_address_on_invalid_handle_is_invalid_state() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]udp-socket.get-remote-address",
            vec![Value::Null]
        )),
        Some("invalid-state".to_string())
    );
}

#[test]
fn udp_bind_requires_socket_handle() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]udp-socket.bind",
            vec![Value::Null, s("127.0.0.1:0")]
        )),
        Some("invalid-state".to_string())
    );
}

#[test]
fn udp_bind_with_missing_address_is_invalid_argument() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]udp-socket.bind",
            vec![udp_socket()]
        )),
        Some("invalid-argument".to_string())
    );
}

#[test]
fn udp_send_requires_socket_handle() {
    assert_eq!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]udp-socket.send",
            vec![Value::Null, Value::Null, Value::Null]
        )),
        Some("invalid-state".to_string())
    );
}

/// 0.2 reached the datagram path through `udp.stream`, which handed back an
/// `outgoing-datagram-stream` resource. 0.3.1 sends on the socket itself, and
/// an unconnected socket with no explicit remote has nowhere to send.
#[test]
fn udp_send_without_a_remote_on_an_unconnected_socket_fails() {
    let socket = udp_socket();
    let bound = call_import(
        "wasi:sockets/types",
        "[method]udp-socket.bind",
        vec![socket.clone(), s(&format!("127.0.0.1:{}", free_port()))],
    );
    assert!(is_ok(&bound), "bind failed: {:?}", err_code(&bound));

    let sent = call_import(
        "wasi:sockets/types",
        "[method]udp-socket.send",
        vec![socket.clone(), Value::Null, Value::Null],
    );
    assert!(
        err_code(&sent).is_some(),
        "an unconnected socket with no remote address cannot send"
    );

    // …and the socket still has no remote address to report.
    assert!(
        err_code(&call_import(
            "wasi:sockets/types",
            "[method]udp-socket.get-remote-address",
            vec![socket]
        ))
        .is_some(),
        "a bound-but-unconnected socket has no remote address"
    );
}

// ── ip-name-lookup ──────────────────────────────────────────────────────────

/// 0.3.1 answers `list<ip-address>` DIRECTLY. 0.2 answered a
/// `resolve-address-stream` resource, and this tree modelled that as an object
/// carrying `__addresses`; asserting the list shape is what stops that
/// resource creeping back.
#[test]
fn resolve_addresses_answers_a_list_not_a_stream_resource() {
    let resolved = call_import(
        "wasi:sockets/ip-name-lookup",
        "resolve-addresses",
        vec![s("localhost")],
    );
    let Value::Object(object) = &resolved else {
        panic!("resolve-addresses must answer a list, got {resolved:?}");
    };
    let object = object.lock().unwrap();
    assert!(
        matches!(&object.kind, ObjectKind::Array(_)),
        "resolve-addresses must answer list<ip-address>, not a resource handle"
    );
    assert!(
        object.properties.get("__addresses").is_none(),
        "the 0.2 `resolve-address-stream` resource must not come back"
    );
}

#[test]
fn resolve_addresses_answers_an_empty_list_for_an_unknown_host() {
    let resolved = call_import(
        "wasi:sockets/ip-name-lookup",
        "resolve-addresses",
        vec![s("definitely.invalid.vybe.test")],
    );
    assert_eq!(len_of(&resolved), 0);
}

/// The `network` handle 0.2 took as `resolve-addresses`' first argument went
/// with the `instance-network` interface. The name is the only argument now.
#[test]
fn resolve_addresses_takes_the_name_as_its_only_argument() {
    assert!(
        len_of(&call_import(
            "wasi:sockets/ip-name-lookup",
            "resolve-addresses",
            vec![s("localhost")]
        )) > 0,
        "localhost must resolve when the name is passed alone"
    );
}
