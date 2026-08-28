//! .NET `System.Net` HTTP adapter — bytecode-only.
//!
//! Lowers the one-call .NET shapes (`WebClient.DownloadString(url)`,
//! `HttpClient.GetStringAsync(url)`, `WebRequest.Create(url).GetResponse()`)
//! onto the **real WASI 0.3 HTTP interfaces**
//! (`wasi:http@0.3.0-rc-2025-09-16`). There is no `fetch`-shaped WASI function
//! — the surface is resource-based, so one .NET call expands to the full
//! request → send → consume-body sequence:
//!
//! 1. `wasi:http/types.[constructor]fields`         → headers handle
//! 2. `wasi:http/types.[static]request.new`          → request handle
//! 3. `[method]request.set-{method,scheme,authority,path-with-query}`
//! 4. `wasi:http/client.send(request)`               → response
//! 5. `wasi:http/types.[static]response.consume-body`→ body stream
//! 6. `wasi:io/streams.[method]input-stream.blocking-read` → bytes
//!
//! This replaces the previous `HostTarget::new("wasi:http", "fetch")` mapping,
//! which named a module nobody registers: `wasi:http` is a WASI *package*, not
//! an interface (0.3 defines `wasi:http/types`, `client` and `handler`), so
//! every `System.Net` call used to emit an unresolvable import.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

fn call_import(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

/// `str.indexOf(needle)` → i32. Stack: `[]` → `[idx]` (reads `s` from a local).
fn index_of(chunk: &mut Chunk, s: u16, needle: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const(needle, line);
    call_import(chunk, "ecma:string", "indexOf", 2, line);
}

/// `str.substring(start[, end])` from locals. Stack: `[]` → `[substring]`.
fn substring(chunk: &mut Chunk, s: u16, start: Src, end: Option<Src>, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    start.emit(chunk, line);
    match end {
        Some(e) => {
            e.emit(chunk, line);
            call_import(chunk, "ecma:string", "substring", 3, line);
        }
        None => call_import(chunk, "ecma:string", "substring", 2, line),
    }
}

/// A substring bound: either a constant or a local slot (optionally offset).
#[derive(Clone, Copy)]
enum Src {
    Const(i32),
    Local(u16),
    LocalPlus(u16, i32),
}

impl Src {
    fn emit(self, chunk: &mut Chunk, line: u32) {
        match self {
            Src::Const(v) => chunk.emit_i32_const(v, line),
            Src::Local(slot) => chunk.emit_op_u16(Op::LOCAL_GET, slot, line),
            Src::LocalPlus(slot, delta) => {
                chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
                chunk.emit_i32_const(delta, line);
                chunk.emit_op(Op::I32_ADD, line);
            }
        }
    }
}

/// Lower a one-call .NET fetch onto the WASI request sequence.
/// Stack: `[url]` → `[body_string]`.
///
/// The URL is split with `ecma:string` primitives (no new host function):
/// `scheme` = text before `"://"` (default `"http"`), `authority` = up to the
/// next `/`, `path` = the remainder (default `"/"`).
pub fn emit_http_fetch(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];

    let base = chunk.local_count;
    chunk.alloc_scratch(8);
    let (url, rest, idx) = (base, base + 1, base + 2);
    let (scheme, authority, path) = (base + 3, base + 4, base + 5);
    let (request, response) = (base + 6, base + 7);

    chunk.emit_op_u16(Op::LOCAL_SET, url, line);

    // ── scheme / rest ───────────────────────────────────────────────────
    index_of(chunk, url, "://", line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_if(line);
    // no scheme — default to http, the whole string is the remainder
    chunk.emit_string_const("http", line);
    chunk.emit_op_u16(Op::LOCAL_SET, scheme, line);
    chunk.emit_op_u16(Op::LOCAL_GET, url, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rest, line);
    chunk.emit_else(line);
    substring(chunk, url, Src::Const(0), Some(Src::Local(idx)), line);
    chunk.emit_op_u16(Op::LOCAL_SET, scheme, line);
    substring(chunk, url, Src::LocalPlus(idx, 3), None, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rest, line);
    chunk.emit_end(line);

    // ── authority / path ────────────────────────────────────────────────
    index_of(chunk, rest, "/", line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_if(line);
    // no path component — whole remainder is the authority
    chunk.emit_op_u16(Op::LOCAL_GET, rest, line);
    chunk.emit_op_u16(Op::LOCAL_SET, authority, line);
    chunk.emit_string_const("/", line);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);
    chunk.emit_else(line);
    substring(chunk, rest, Src::Const(0), Some(Src::Local(idx)), line);
    chunk.emit_op_u16(Op::LOCAL_SET, authority, line);
    substring(chunk, rest, Src::Local(idx), None, line);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);
    chunk.emit_end(line);

    // ── request = request.new(headers, contents, trailers, options) ─────
    // WASI 0.3 static constructor (`wasi:http@0.3.0-rc-2025-09-16`). A GET has
    // no body, so `contents` / `trailers` / `options` are all absent.
    call_import(chunk, "wasi:http/types", "[constructor]fields", 0, line);
    core_wasm::null(chunk, line); // contents:  option<stream<u8>>
    core_wasm::null(chunk, line); // trailers:  future<...>
    core_wasm::null(chunk, line); // options:   option<request-options>
    call_import(chunk, "wasi:http/types", "[static]request.new", 4, line);
    chunk.emit_op_u16(Op::LOCAL_SET, request, line);

    // set-method / set-scheme / set-authority / set-path-with-query
    chunk.emit_op_u16(Op::LOCAL_GET, request, line);
    chunk.emit_string_const("GET", line);
    call_import(
        chunk,
        "wasi:http/types",
        "[method]request.set-method",
        2,
        line,
    );
    chunk.emit_op(Op::DROP, line);
    for (setter, slot) in [
        ("[method]request.set-scheme", scheme),
        ("[method]request.set-authority", authority),
        ("[method]request.set-path-with-query", path),
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, request, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        call_import(chunk, "wasi:http/types", setter, 2, line);
        chunk.emit_op(Op::DROP, line);
    }

    // ── response = client.send(request) ─────────────────────────────────
    // WASI 0.3 `interface client { send: async func(request) -> result<
    // response, error-code> }` — one call, the response comes back directly
    // (0.2 needed outgoing-handler.handle -> future-incoming-response -> get).
    chunk.emit_op_u16(Op::LOCAL_GET, request, line);
    call_import(chunk, "wasi:http/client", "send", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, response, line);

    // ── body ────────────────────────────────────────────────────────────
    // `response.consume-body(this, res)` (0.3 static) yields
    // `tuple<stream<u8>, future<...>>`; the shared read drains the stream.
    //
    // This used to call `wasi:io/streams.[method]input-stream.blocking-read`
    // on it — a WASI **0.2** method applied to a **0.3** stream, against a
    // `wasi:io` package 0.3 DELETED. It only ever resolved because the host
    // still registers a 0.2.12 compat provider.
    chunk.emit_op_u16(Op::LOCAL_GET, response, line);
    core_wasm::null(chunk, line); // res: future<result<_, error-code>>
    call_import(
        chunk,
        "wasi:http/types",
        "[static]response.consume-body",
        2,
        line,
    );
    // Element 0 of the tuple is the `stream<u8>` — an i32 readable handle,
    // per the canonical ABI.
    chunk.emit_i32_const(0, line);
    call_import(chunk, "ecma:array", "at", 2, line);
    vybe_compiler::primitives::io::emit_read_stream_to_string(chunk, line);
}

/// `WebRequest.Create(url)` / `new WebClient()` / `new HttpClient()` — these
/// .NET objects carry no state here (each request is self-contained), so
/// construction yields a typed marker. Stack: `[..args]` → `[obj]`.
pub fn emit_http_client_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    class_slots::emit_class_alloc(chunk, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const("HttpClient", line);
    let idx = chunk.add_constant(Value::String(Arc::from("__type")));
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot("__type"),
        ValueSource::Stack,
        line,
    );
}
