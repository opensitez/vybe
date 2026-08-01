//! `node:http.*` — per-request read access.
//!
//! All accessors return sentinel values (empty string, null, 0, empty
//! array) when no request context is installed on the current thread
//! ("cli" mode). This makes the primitives safe to call from any script,
//! whether it runs under `--serve` or at the terminal.
//!
//! Parsed accessors (cookies, query pairs, content-type, form bodies,
//! auth) will land in a subsequent slice — they cache their result on the
//! `RequestContext` via `OnceLock`.

use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

use super::context::with_context;

pub fn register(vm: &mut VM) {
    // ── IncomingMessage as a Readable ──────────────────────────────────────
    //
    // Node reads a request body by treating `IncomingMessage` as a stream:
    // `readable.read([size])` returns the next chunk, or `null` at end of
    // stream (Node docs, `stream.Readable.read`). That null-at-EOF contract is
    // what tells a consumer to stop; a sentinel empty string would loop.
    vm.register_host_fn(
        "node:http",
        "read",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let size = match args.first() {
                Some(Value::F64(n)) if *n > 0.0 => *n as usize,
                Some(Value::I32(n)) if *n > 0 => *n as usize,
                _ => 64 * 1024,
            };
            with_context(|c| {
                let mut body = c.body.lock().unwrap();
                if body.eof() {
                    return Value::Null;
                }
                let bytes = body.read(size);
                if bytes.is_empty() {
                    return Value::Null;
                }
                let elems: Vec<Value> = bytes.into_iter().map(|b| Value::I32(b as i32)).collect();
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems)))
            })
            .unwrap_or(Value::Null)
        }),
    );

    // `readable.readableEnded` — true once the stream is exhausted.
    vm.register_host_fn(
        "node:http",
        "readable_ended",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            with_context(|c| Value::Bool(c.body.lock().unwrap().eof())).unwrap_or(Value::Bool(true))
        }),
    );

    // `message.httpVersion` — the version the request came in on.
    vm.register_host_fn(
        "node:http",
        "http_version",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::String(Arc::from("1.1"))),
    );

    // Raw accessors ────────────────────────────────────────────────────────
    vm.register_host_fn(
        "node:http",
        "method",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            with_context(|c| Value::String(Arc::from(c.method.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "uri",
        Box::new(|_ctx, _| {
            with_context(|c| Value::String(Arc::from(c.uri.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );







    vm.register_host_fn(
        "node:http",
        "remote_addr",
        Box::new(|_ctx, _| {
            with_context(|c| Value::String(Arc::from(c.remote_addr.as_str())))
                .unwrap_or_else(|| Value::String(Arc::from("")))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "remote_port",
        Box::new(|_ctx, _| {
            with_context(|c| Value::F64(c.remote_port as f64)).unwrap_or(Value::F64(0.0))
        }),
    );

    // Headers ──────────────────────────────────────────────────────────────
    vm.register_host_fn(
        "node:http",
        "header",
        Box::new(|_ctx, args| {
            let name = string_arg(args, 0).to_ascii_lowercase();
            with_context(|c| {
                for (n, v) in &c.headers {
                    if n.eq_ignore_ascii_case(&name) {
                        return Value::String(Arc::from(v.as_str()));
                    }
                }
                Value::Null
            })
            .unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        "node:http",
        "header_all",
        Box::new(|_ctx, args| {
            let name = string_arg(args, 0).to_ascii_lowercase();
            with_context(|c| {
                let mut out = Vec::new();
                for (n, v) in &c.headers {
                    if n.eq_ignore_ascii_case(&name) {
                        out.push(Value::String(Arc::from(v.as_str())));
                    }
                }
                array_value(out)
            })
            .unwrap_or_else(|| array_value(Vec::new()))
        }),
    );

    vm.register_host_fn(
        "node:http",
        "headers",
        Box::new(|_ctx, _| {
            with_context(|c| {
                let items = c.headers.iter().map(|(n, v)| pair_object(n, v)).collect();
                array_value(items)
            })
            .unwrap_or_else(|| array_value(Vec::new()))
        }),
    );









}



// Helpers ──────────────────────────────────────────────────────────────────

fn string_arg(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}

fn array_value(items: Vec<Value>) -> Value {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Array(items);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn pair_object(name: &str, value: &str) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("name".into(), Value::String(Arc::from(name)));
    obj.properties
        .insert("value".into(), Value::String(Arc::from(value)));
    Value::Object(vybe_runtime::heap::alloc(obj))
}
