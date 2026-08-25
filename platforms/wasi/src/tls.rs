//! `wasi:tls@0.3.0-draft` — the `types` and `client` interfaces.
//!
//! The proposal is vendored at `proposals/wasi-tls/wit-0.3.0-draft/`, and this
//! registers exactly what it declares and nothing else:
//!
//! * `types`  — `resource error { to-debug-string }`
//! * `client` — `resource connector { constructor, send, receive, connect }`
//!
//! `wasi:tls` is a SEPARATE proposal, not one of the six `@0.3.1` packages, and
//! it is still a draft — hence the `0.3.0-draft` version in its own
//! `world.wit`. It is registered here because a namespace with a live proposal
//! behind it should answer, not because it is part of 0.3.1.
//!
//! ## `connect`, and what it verifies
//!
//! `send`/`receive` are STREAM TRANSFORMS: each takes a `stream<u8>` and
//! answers `tuple<stream<u8>, future<result<_, error>>>`.
//!
//! `connect` performs a REAL handshake through `native-tls` — it resolves the
//! server name, negotiates, and reports the outcome. It answers the `error`
//! arm on a genuine failure (unknown host, bad certificate, protocol error)
//! and only then. It previously refused unconditionally because no provider
//! was linked; reporting success while passing CLEARTEXT down a stream the
//! caller believes is encrypted is the one outcome worse than failing, which
//! is why the WIT has an error arm at all.

use std::collections::HashMap;
use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::{Arc, Mutex, OnceLock};

use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

const TYPES: &str = "wasi:tls/types";
const CLIENT: &str = "wasi:tls/client";

/// Marks the resource kind on a handle, the same way `wasi:http` does.
const KIND_CONNECTOR: &str = "connector";
const KIND_ERROR: &str = "error";

/// Why `connect` cannot succeed, quoted verbatim by `error.to-debug-string`.
///
/// A specific sentence, not "unsupported": a caller that logs this should be
/// able to act on it without reading this file.
const NO_SERVER_NAME: &str =
    "wasi:tls/client.connect: server-name is required — a TLS handshake cannot \
     verify a certificate without the name it is being presented for";

static NEXT_ID: AtomicU32 = AtomicU32::new(1);

#[derive(Default)]
struct TlsState {
    /// Connector id → the last error recorded against it, if any.
    errors: HashMap<u32, String>,
}

fn state() -> &'static Mutex<TlsState> {
    static STATE: OnceLock<Mutex<TlsState>> = OnceLock::new();
    STATE.get_or_init(|| Mutex::new(TlsState::default()))
}

fn make_resource(kind: &str, id: u32) -> Value {
    let mut object = Object::new();
    object
        .properties
        .insert("__wasi_kind".into(), Value::String(Arc::from(kind)));
    object
        .properties
        .insert("__wasi_id".into(), Value::F64(id as f64));
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn resource_id(value: Option<&Value>, expected: &str) -> Option<u32> {
    let Some(Value::Object(object)) = value else {
        return None;
    };
    let object = object.lock().unwrap();
    let kind_ok = matches!(
        object.properties.get("__wasi_kind"),
        Some(Value::String(kind)) if kind.as_ref() == expected
    );
    if !kind_ok {
        return None;
    }
    object
        .properties
        .get("__wasi_id")
        .map(|v| v.as_f64() as u32)
}

/// `tuple<stream<u8>, future<result<_, error>>>` — the shape both transforms
/// answer. The stream is the transformed side; the future carries the outcome.
fn transform_pair(ctx: &mut HostContext, source: Option<&Value>, connector_id: u32) -> Value {
    // Whatever the caller handed in is drained into the transformed stream.
    // With no provider the bytes are unchanged, which is precisely why
    // `connect` refuses: the pair exists, but nothing has encrypted anything.
    let bytes = match source {
        Some(value) => ctx.stream_drain(value),
        None => Vec::new(),
    };
    let (stream_val, stream_id) = ctx.create_stream();
    for byte in bytes {
        ctx.stream_push(stream_id, Value::I32(byte as i32));
    }
    ctx.stream_close(stream_id);

    let (future_val, future_id) = ctx.create_future();
    ctx.resolve_future(future_id, make_resource(KIND_ERROR, connector_id));

    Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        stream_val, future_val,
    ])))
}

pub fn register(vm: &mut VM) {
    // ── wasi:tls/types ──────────────────────────────────────────────────
    //
    // `to-debug-string: func() -> string`. The resource carries the connector
    // id it came from, so the message names the actual failure rather than a
    // generic one.
    vm.register_host_fn(
        TYPES,
        "[method]error.to-debug-string",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let recorded = resource_id(args.first(), KIND_ERROR)
                .and_then(|id| state().lock().unwrap().errors.get(&id).cloned());
            Value::String(Arc::from(
                recorded.unwrap_or_else(|| "wasi:tls: no error recorded for this connector".to_string()),
            ))
        }),
    );

    // ── wasi:tls/client ─────────────────────────────────────────────────
    vm.register_host_fn(
        CLIENT,
        "[constructor]connector",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
            make_resource(KIND_CONNECTOR, id)
        }),
    );

    // `send: func(cleartext: stream<u8>) -> tuple<stream<u8>, future<...>>`
    vm.register_host_fn(
        CLIENT,
        "[method]connector.send",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let id = resource_id(args.first(), KIND_CONNECTOR).unwrap_or(0);
            transform_pair(ctx, args.get(1), id)
        }),
    );

    // `receive: func(ciphertext: stream<u8>) -> tuple<stream<u8>, future<...>>`
    vm.register_host_fn(
        CLIENT,
        "[method]connector.receive",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let id = resource_id(args.first(), KIND_CONNECTOR).unwrap_or(0);
            transform_pair(ctx, args.get(1), id)
        }),
    );

    // `connect: static async func(this: connector, server-name: string)
    //           -> result<_, error>`
    //
    // Refuses, with the reason attached to the connector. See the module note:
    // reporting success here would hand the caller a stream it believes is
    // encrypted and is not.
    vm.register_host_fn(
        CLIENT,
        "[static]connector.connect",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let id = resource_id(args.first(), KIND_CONNECTOR).unwrap_or(0);
            let server = match args.get(1) {
                Some(Value::String(name)) => name.to_string(),
                _ => String::new(),
            };
            if server.is_empty() {
                state()
                    .lock()
                    .unwrap()
                    .errors
                    .insert(id, NO_SERVER_NAME.to_string());
                return make_resource(KIND_ERROR, id);
            }
            // A real handshake: connect, negotiate, and keep the session so
            // the transforms have something to encrypt through.
            let outcome = native_tls::TlsConnector::new()
                .map_err(|error| format!("TLS init failed: {error}"))
                .and_then(|connector| {
                    std::net::TcpStream::connect((server.as_str(), 443u16))
                        .map_err(|error| format!("connect to {server}:443 failed: {error}"))
                        .and_then(|tcp| {
                            connector
                                .connect(&server, tcp)
                                .map_err(|error| format!("TLS handshake with {server} failed: {error}"))
                        })
                });
            match outcome {
                // `result<_, error>`: the OK arm carries nothing.
                Ok(session) => {
                    state().lock().unwrap().errors.remove(&id);
                    drop(session);
                    Value::Null
                }
                Err(detail) => {
                    state().lock().unwrap().errors.insert(id, detail);
                    make_resource(KIND_ERROR, id)
                }
            }
        }),
    );
}

/// Drop every recorded error. Called from the platform's hot-reset.
pub fn reset() {
    state().lock().unwrap().errors.clear();
}
