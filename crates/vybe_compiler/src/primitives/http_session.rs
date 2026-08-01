//! Sessions — identity and lifecycle, shared across languages.
//!
//! No spec surface has sessions: `wasi:http` carries messages, Node's session
//! support is userland (`express-session`), and there is no `wasi:keyvalue` in
//! this repo. So per `documentation/httpserver.md` §4a this is a PRIMITIVE,
//! emitted once over the spec surfaces, rather than a host function or — as it
//! was — Rust inside the vybex server that only PHP could reach.
//!
//! One implementation backs PHP `session_*`/`$_SESSION`, Python's session
//! mappings, Rack's `session` and ASP.NET's `Session`. That matters because a
//! session is a session ACROSS languages: PHP can call C# can call Python in
//! one process, and all three have to see the same session.
//!
//! **The cookie name is a parameter, never a constant here.** PHP spells it
//! `PHPSESSID`, ASP.NET `ASP.NET_SessionId`, Flask `session`. The language
//! adapter passes its own spelling in; putting a table of them in this file
//! would be `php_lang.rs` all over again.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// This request's session id. Published by the host when it already knows the
/// id, otherwise generated on first use.
pub const SESSION_ID_GLOBAL: &str = "__vybe_session_id";

/// The session's data map — what `$_SESSION` and friends alias.
pub const SESSION_DATA_GLOBAL: &str = "__vybe_session_data";

/// Tri-state lifecycle, matching the values every language already uses:
/// 0 disabled, 1 not started, 2 active.
pub const SESSION_STATUS_GLOBAL: &str = "__vybe_session_status";

/// The cookie name in force for this request, once `start` has run.
pub const SESSION_NAME_GLOBAL: &str = "__vybe_session_name";

/// Not started — no session has been opened on this request.
pub const STATUS_NONE: i32 = 1;
/// Active — a session is open.
pub const STATUS_ACTIVE: i32 = 2;

fn push_global(chunk: &mut Chunk, name: &str, line: u32) {
    let key = chunk.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(name)));
    chunk.emit_op_u16(Op::GLOBAL_GET, key, line);
}

fn set_global(chunk: &mut Chunk, name: &str, line: u32) {
    let key = chunk.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(name)));
    chunk.emit_op_u16(Op::GLOBAL_SET, key, line);
}

/// The session cookie name. Stack: [] → [string].
///
/// `default_name` is the calling language's spelling, used until something
/// overrides it (PHP's `session_name($new)`).
pub fn emit_name(chunks: &mut [Chunk], current: usize, default_name: &str, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    push_global(&mut chunks[current], SESSION_NAME_GLOBAL, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    push_global(&mut chunks[current], SESSION_NAME_GLOBAL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const(default_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Set the session cookie name. Stack: [string] → [].
pub fn emit_set_name(chunks: &mut [Chunk], current: usize, line: u32) {
    set_global(&mut chunks[current], SESSION_NAME_GLOBAL, line);
}

/// The session lifecycle state. Stack: [] → [i32]. Defaults to "not started".
pub fn emit_status(chunks: &mut [Chunk], current: usize, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    push_global(&mut chunks[current], SESSION_STATUS_GLOBAL, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    push_global(&mut chunks[current], SESSION_STATUS_GLOBAL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(STATUS_NONE, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// This request's session id. Stack: [] → [string].
///
/// Resolution order: an id the host already published, then the request's
/// session cookie, then a freshly generated one. Memoised, so every language in
/// the process agrees on the id.
///
/// The id is emitted RAW. Real PHP does not percent-encode session ids — `-`
/// and hex digits are all valid cookie-octets — and an encoded id would come
/// back out of `session_id()` encoded, which is the value apps put in URLs.
pub fn emit_id(chunks: &mut [Chunk], current: usize, default_name: &str, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    let from_cookie = chunks[current].alloc_scratch(1);

    push_global(&mut chunks[current], SESSION_ID_GLOBAL, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    push_global(&mut chunks[current], SESSION_ID_GLOBAL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_else(line);

    // The client's cookie, if it sent one.
    super::http_cookie::emit_request_cookies(chunks, current, line);
    emit_name(chunks, current, default_name, line);
    super::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, from_cookie, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, from_cookie, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_cookie, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_else(line);
    emit_new_id(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    set_global(&mut chunks[current], SESSION_ID_GLOBAL, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Set the session id explicitly. Stack: [string] → [].
pub fn emit_set_id(chunks: &mut [Chunk], current: usize, line: u32) {
    set_global(&mut chunks[current], SESSION_ID_GLOBAL, line);
}

/// A fresh session id: 32 lowercase hex chars. Stack: [] → [string].
///
/// Bytes come from `wasi:random/random.get-random-bytes`, which the spec
/// defines as cryptographically strong — a session id is a bearer credential,
/// so the insecure generator would be a real vulnerability, not a style choice.
pub fn emit_new_id(chunks: &mut [Chunk], current: usize, line: u32) {
    let bytes = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);

    chunks[current].emit_i32_const(16, line);
    let get_bytes = chunks[current].add_import("wasi:random/random", "get-random-bytes");
    chunks[current].emit_call(get_bytes, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bytes, line);

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    // Two hex digits per byte, zero-padded — a 31-char id would be a subtly
    // different credential each time a byte happened to be < 16.
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    emit_hex_nibble(chunks, current, bytes, i, true, line);
    emit_hex_nibble(chunks, current, bytes, i, false, line);
    super::strings::emit_concat(&mut chunks[current], 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// One hex digit of `bytes[i]` — the high nibble when `high`, else the low one.
/// Stack: [] → [string].
fn emit_hex_nibble(
    chunks: &mut [Chunk],
    current: usize,
    bytes: u16,
    i: u16,
    high: bool,
    line: u32,
) {
    let nibble = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    if high {
        chunks[current].emit_i32_const(4, line);
        chunks[current].emit_op(Op::I32_SHR_U, line);
    }
    chunks[current].emit_i32_const(15, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nibble, line);

    // `substring(s, n, n+1)` — the js-string builtins have no `charAt`.
    chunks[current].emit_string_const("0123456789abcdef", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nibble, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nibble, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    let substring = chunks[current].add_import("wasm:js-string", "substring");
    chunks[current].emit_call(substring, 3, line);
}

/// The session's data map. Stack: [] → [map].
///
/// This is the object a language aliases as `$_SESSION` / `session` / `Session`.
/// Memoised under one global so every language in the process mutates the SAME
/// map — that is what makes a session survive a PHP → C# → Python call chain.
pub fn emit_data(chunks: &mut [Chunk], current: usize, line: u32) {
    super::http_request_env::emit_memoised(
        chunks,
        current,
        SESSION_DATA_GLOBAL,
        line,
        |chunks, current, line| {
            super::collections::emit_map_new(chunks, current, line);
        },
    );
}

/// Open the session. Stack: [] → [bool].
///
/// Idempotent: starting an already-active session is a no-op that reports
/// success, which is what every language does when a framework and user code
/// both call it.
pub fn emit_start(chunks: &mut [Chunk], current: usize, default_name: &str, line: u32) {
    // Resolving the id is what adopts the client's cookie or mints a new one.
    emit_id(chunks, current, default_name, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_data(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(STATUS_ACTIVE, line);
    set_global(&mut chunks[current], SESSION_STATUS_GLOBAL, line);
    chunks[current].emit_i32_const(1, line);
    super::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Replace the session id, keeping the data. Stack: [] → [bool].
///
/// The defence against session fixation: after a privilege change the client
/// gets a new id it could not have chosen.
pub fn emit_regenerate_id(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_new_id(chunks, current, line);
    set_global(&mut chunks[current], SESSION_ID_GLOBAL, line);
    chunks[current].emit_i32_const(1, line);
    super::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Discard the session's contents and close it. Stack: [] → [bool].
pub fn emit_destroy(chunks: &mut [Chunk], current: usize, line: u32) {
    super::collections::emit_map_new(chunks, current, line);
    set_global(&mut chunks[current], SESSION_DATA_GLOBAL, line);
    chunks[current].emit_i32_const(STATUS_NONE, line);
    set_global(&mut chunks[current], SESSION_STATUS_GLOBAL, line);
    chunks[current].emit_i32_const(1, line);
    super::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Empty the session's data without closing it. Stack: [] → [].
pub fn emit_unset(chunks: &mut [Chunk], current: usize, line: u32) {
    super::collections::emit_map_new(chunks, current, line);
    set_global(&mut chunks[current], SESSION_DATA_GLOBAL, line);
}
