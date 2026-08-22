//! The request, as a language-neutral environment.
//!
//! PHP's `$_SERVER`/`$_GET`, WSGI's `environ`, Rack's `env`, ASP.NET's
//! `HttpContext.Request` and Node's `req` are the SAME DATA in five spellings.
//! This primitive produces the data once; a language adapter renames it.
//! See `documentation/httpserver.md` §4a.
//!
//! Everything here is composed from `wasi:http` — the spec HTTP surface — read
//! through the handle the server publishes as a global. Nothing in this file
//! knows about any language, and nothing about HTTP is re-implemented: the
//! parsing that no spec covers (cookies, multipart) lives in its own primitive.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Global holding this request's `wasi:http` `request` handle.
///
/// The spec has the host hand a `request` to a component's exported
/// `handler.handle`. `vybex --serve` compiles a SCRIPT, which has no exports,
/// so the handle is published under this reserved name instead. Kept in step
/// with `crates/vybex/src/server/script.rs`.
///
/// 0.3.1 has ONE `request` resource; 0.2's `incoming-request` was the arriving
/// half of a pair. The global keeps its 0.2-flavoured NAME because renaming a
/// reserved global is a separate, cross-crate change — the WASI names it is
/// read with are the 0.3.1 ones.
pub const REQUEST_GLOBAL: &str = "__wasi_http_incoming_request";

/// Global holding the `response` this request answers with.
///
/// 0.3.1 has `handler.handle` RETURN a `result<response, error-code>`. 0.2
/// instead handed the guest a `response-outparam` to write into, and
/// `wasi:http@0.3.1` does not declare that resource at all — there is nothing
/// left to call `set` on.
///
/// A script has no export to return FROM, so the resource is assigned to this
/// reserved global and the host takes delivery from it: the same substitution
/// [`REQUEST_GLOBAL`] already makes in the other direction, and the reason both
/// are reserved names rather than invented `wasi:` functions.
pub const RESPONSE_GLOBAL: &str = "__wasi_http_response";

/// Global holding DEPLOYMENT metadata the transport supplies, CGI-named.
///
/// `wasi:http` models the message, not the deployment: it has no document
/// root, script path, server software string, peer address or protocol
/// version. Those come from the server, under their standard CGI names so the
/// map stays language-neutral — no adapter renames them.
pub const SERVER_ENV_GLOBAL: &str = "__wasi_http_server_env";

/// Memoised result of [`emit_environ`], so every language sees ONE map.
///
/// Identity matters, not just shape: a process can run PHP calling C# calling
/// Python, and `$_SERVER`, `HttpContext` and WSGI `environ` must be the same
/// object or a write through one name is invisible through another.
const ENVIRON_CACHE_GLOBAL: &str = "__vybe_request_environ";

/// The request body, read once and kept.
const BODY_CACHE_GLOBAL: &str = "__vybe_request_body";

fn push_request_handle(chunk: &mut Chunk, line: u32) {
    crate::primitives::globals::emit_read(chunk, REQUEST_GLOBAL, line);
}

/// Call a `wasi:http/types` method on the request handle.
/// Stack: [] → [result].
fn emit_request_method_call(chunks: &mut [Chunk], current: usize, method: &str, line: u32) {
    push_request_handle(&mut chunks[current], line);
    let idx = chunks[current].add_import("wasi:http/types", method);
    chunks[current].emit_call(idx, 1, line);
}

/// The HTTP method (`GET`, `POST`, …). Stack: [] → [string].
pub fn emit_method(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_request_method_call(chunks, current, "[method]request.get-method", line);
}

/// Path AND query, exactly as `wasi:http` reports it. Stack: [] → [string|null].
pub fn emit_path_with_query(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_request_method_call(
        chunks,
        current,
        "[method]request.get-path-with-query",
        line,
    );
}

/// URI scheme, or null. Stack: [] → [string|null].
pub fn emit_scheme(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_request_method_call(chunks, current, "[method]request.get-scheme", line);
}

/// URI authority (`host[:port]`), or null. Stack: [] → [string|null].
pub fn emit_authority(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_request_method_call(chunks, current, "[method]request.get-authority", line);
}

/// The request headers as a `wasi:http` `fields` resource. Stack: [] → [fields].
pub fn emit_headers(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_request_method_call(chunks, current, "[method]request.get-headers", line);
}

/// The path with any query string removed. Stack: [] → [string].
///
/// `wasi:http` reports one `path-with-query`; most language surfaces want the
/// two halves separately (`PATH_INFO` / `QUERY_STRING`, `$_SERVER` vs `$_GET`).
pub fn emit_path(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_path_with_query(chunks, current, line);
    emit_split_on_question_mark(chunks, current, 0, line);
}

/// The query string alone, empty when there is none. Stack: [] → [string].
pub fn emit_query_string(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_path_with_query(chunks, current, line);
    emit_split_on_question_mark(chunks, current, 1, line);
}

/// Split `[string]` on the first `?` and keep `part` (0 = before, 1 = after).
/// Stack: [string] → [string]. A missing `?` yields the whole string for part 0
/// and an empty string for part 1.
fn emit_split_on_question_mark(chunks: &mut [Chunk], current: usize, part: u8, line: u32) {
    let src = chunks[current].alloc_scratch(1);
    let idx_slot = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    // A null path-with-query behaves as "".
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_string_const("?", line);
    let index_of = chunks[current].add_import("ecma:string", "indexOf");
    chunks[current].emit_call(index_of, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    // No '?': the whole string is the path; the query is empty.
    if part == 0 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    if part == 0 {
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
        let len = chunks[current].add_import("wasm:js-string", "length");
        chunks[current].emit_call(len, 1, line);
    }
    let substring = chunks[current].add_import("wasm:js-string", "substring");
    chunks[current].emit_call(substring, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// One request header value by (case-insensitive) name, or null.
/// Stack: [] → [string|null].
///
/// `wasi:http` `fields.get` returns a LIST of values (a header may repeat);
/// callers that want a single value take the first, which is what every
/// CGI-shaped surface reports.
pub fn emit_header(chunks: &mut [Chunk], current: usize, name: &str, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    emit_headers(chunks, current, line);
    chunks[current].emit_string_const(name, line);
    let get = chunks[current].add_import("wasi:http/types", "[method]fields.get");
    chunks[current].emit_call(get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, list, line);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    // `wasi:http` header values are `list<u8>` (§fields), so the first entry
    // is a BYTE ARRAY, not a string.
    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    emit_bytes_to_string(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

fn push_global(chunk: &mut Chunk, name: &str, line: u32) {
    crate::primitives::globals::emit_read(chunk, name, line);
}

fn set_global(chunk: &mut Chunk, name: &str, line: u32) {
    crate::primitives::globals::emit_write(chunk, name, line);
}

/// Store `[value]` into the environ map under `key`. Stack: [map, value] → [map].
fn emit_put(chunks: &mut [Chunk], current: usize, map: u16, key: &str, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// The request as a CGI-shaped environment. Stack: [] → [map].
///
/// CGI is the SHARED spelling, not PHP's: WSGI `environ` is CGI keys plus
/// `wsgi.*`, Rack `env` is CGI keys plus `rack.*`, and PHP's `$_SERVER` is the
/// bare map. So this produces the common set and each language adds its own
/// few extras — nothing here renames anything for anyone.
///
/// Built ONCE per request and memoised (see `ENVIRON_CACHE_GLOBAL`).
pub fn emit_environ(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);

    // Return the cached map if this request already built it.
    push_global(&mut chunks[current], ENVIRON_CACHE_GLOBAL, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    push_global(&mut chunks[current], ENVIRON_CACHE_GLOBAL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map, line);
    chunks[current].emit_else(line);

    // Start from the transport's deployment metadata, then layer the message
    // on top so a value derived from the request always wins.
    push_global(&mut chunks[current], SERVER_ENV_GLOBAL, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map, line);
    chunks[current].emit_else(line);
    push_global(&mut chunks[current], SERVER_ENV_GLOBAL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map, line);
    chunks[current].emit_end(line);

    emit_method(chunks, current, line);
    emit_put(chunks, current, map, "REQUEST_METHOD", line);

    emit_path_with_query(chunks, current, line);
    emit_put(chunks, current, map, "REQUEST_URI", line);

    emit_path(chunks, current, line);
    emit_put(chunks, current, map, "PATH_INFO", line);

    emit_query_string(chunks, current, line);
    emit_put(chunks, current, map, "QUERY_STRING", line);

    // SERVER_NAME / HTTP_HOST both come from the authority.
    emit_authority(chunks, current, line);
    emit_put(chunks, current, map, "HTTP_HOST", line);
    emit_authority(chunks, current, line);
    emit_put(chunks, current, map, "SERVER_NAME", line);

    emit_scheme(chunks, current, line);
    emit_put(chunks, current, map, "REQUEST_SCHEME", line);

    // CGI reports HTTPS as "on" when the scheme is https, and omits it
    // otherwise — Symfony's `Request::isSecure()` tests exactly that.
    let scheme = chunks[current].alloc_scratch(1);
    emit_scheme(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, scheme, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, scheme, line);
    chunks[current].emit_string_const("https", line);
    super::ops::emit_dyn_eq(&mut chunks[current], line);
    super::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("on", line);
    emit_put(chunks, current, map, "HTTPS", line);
    chunks[current].emit_end(line);

    emit_header(chunks, current, "content-type", line);
    emit_put(chunks, current, map, "CONTENT_TYPE", line);
    emit_header(chunks, current, "content-length", line);
    emit_put(chunks, current, map, "CONTENT_LENGTH", line);

    emit_http_header_keys(chunks, current, map, line);

    push_global(&mut chunks[current], ENVIRON_CACHE_GLOBAL, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    set_global(&mut chunks[current], ENVIRON_CACHE_GLOBAL, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
}

/// A byte list to a string. Stack: [array of byte numbers] → [string].
///
/// `wasi:http` hands out header values as `list<u8>`, which arrives as a plain
/// array of numbers. `web:encoding.decode` (WHATWG `TextDecoder`) only accepts
/// an `ArrayBuffer`/`TypedArray`, so it cannot consume this shape — hence the
/// per-byte build over `wasm:js-string.fromCharCode`. Header field values are
/// ISO-8859-1 per RFC 9110 §5.5, so byte-per-char is the correct decoding here
/// (this is NOT a general UTF-8 decoder).
pub fn emit_bytes_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let bytes = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, bytes, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    // A non-array (or absent) value decodes to "".
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

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

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    let from_char = chunks[current].add_import("wasm:js-string", "fromCharCode");
    chunks[current].emit_call(from_char, 1, line);
    super::strings::emit_concat(&mut chunks[current], 2, line);
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

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Add every request header to `map` under its CGI `HTTP_*` name.
///
/// CGI §4.1.18: a header `X-Forwarded-For` becomes `HTTP_X_FORWARDED_FOR` —
/// uppercased with `-` replaced by `_`. Symfony reads these to reconstruct the
/// header bag, and WSGI/Rack use the identical convention, so this is shared
/// rather than PHP-specific.
fn emit_http_header_keys(chunks: &mut [Chunk], current: usize, map: u16, line: u32) {
    let entries = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);

    emit_headers(chunks, current, line);
    let entries_fn = chunks[current].add_import("wasi:http/types", "[method]fields.copy-all");
    chunks[current].emit_call(entries_fn, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
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

    // entries[i] is a [name, value-bytes] pair.
    chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair, line);

    // key = "HTTP_" + name.toUpperCase().replaceAll("-", "_")
    chunks[current].emit_string_const("HTTP_", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    let upper = chunks[current].add_import("ecma:string", "toUpperCase");
    chunks[current].emit_call(upper, 1, line);
    chunks[current].emit_string_const("-", line);
    chunks[current].emit_string_const("_", line);
    let replace_all = chunks[current].add_import("ecma:string", "replaceAll");
    chunks[current].emit_call(replace_all, 3, line);
    super::strings::emit_concat(&mut chunks[current], 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    emit_bytes_to_string(chunks, current, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_end(line);
}

/// The request body as a string. Stack: [] → [string].
///
/// `wasi:http` gives the body up exactly ONCE — `incoming-request.consume`
/// fails on a second call by design (§types, "consumes"). But every language
/// exposes several body-derived surfaces at once: PHP has `$_POST`, `$_FILES`
/// and `php://input`; WSGI has `wsgi.input`; Rack has `rack.input`. Each of
/// those reaching for the body independently would mean the first one wins and
/// the rest silently see nothing. So the read happens once here and the bytes
/// are memoised; every body-derived surface is built from this.
///
/// The bytes become a LATIN-1 string — one char per byte, so the round trip is
/// lossless for binary uploads and the parsers can do ordinary string work.
/// Absent request, failed consume and empty body all give `""`, because no
/// caller wants an error for "this was a GET".
pub fn emit_body(chunks: &mut [Chunk], current: usize, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    let body = chunks[current].alloc_scratch(1);
    let stream = chunks[current].alloc_scratch(1);

    push_global(&mut chunks[current], BODY_CACHE_GLOBAL, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    push_global(&mut chunks[current], BODY_CACHE_GLOBAL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_else(line);

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    push_request_handle(&mut chunks[current], line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    // `[static]request.consume-body(request, trailers)` — the WASI 0.3 shape.
    // This used to be `incoming-request.consume` → `incoming-body.%stream` →
    // `wasi:io/streams.[method]input-stream.blocking-read`, a chain whose last
    // two links 0.3 DELETED: `incoming-body` is gone as a resource and the
    // whole `wasi:io` package with it. 0.3 hands the body over as a
    // `tuple<stream<u8>, future<result<option<trailers>, error-code>>>` in one
    // call, so the stream IS the body and there is no intermediate resource.
    push_request_handle(&mut chunks[current], line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    let consume = chunks[current].add_import("wasi:http/types", "[static]request.consume-body");
    chunks[current].emit_call(consume, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, body, line);

    // A trap-free check: §request.consume-body succeeds at most once, and a
    // second call answers an error record rather than the tuple. Only an
    // actual tuple has a stream at element 0 — asking an error record for one
    // would hand `canon stream.read` a non-handle, which traps.
    chunks[current].emit_op_u16(Op::LOCAL_GET, body, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    chunks[current].emit_if(line);

    // Element 0 of the tuple is the `stream<u8>` — an i32 readable handle,
    // per the canonical ABI.
    chunks[current].emit_op_u16(Op::LOCAL_GET, body, line);
    chunks[current].emit_i32_const(0, line);
    let at = chunks[current].add_import("ecma:array", "at");
    chunks[current].emit_call(at, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stream, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, stream, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    // LATIN-1, deliberately — NOT the UTF-8 `emit_read_stream_to_string`.
    // One char per byte keeps a binary upload lossless through a string, which
    // is what every body parser downstream (`$_POST`, `$_FILES`, `wsgi.input`,
    // `rack.input`) is written against. Decoding as UTF-8 here would replace
    // every byte ≥ 0x80 that is not part of a valid sequence with U+FFFD and
    // silently corrupt every file upload.
    chunks[current].emit_op_u16(Op::LOCAL_GET, stream, line);
    super::io::emit_read_stream_to_bytes(&mut chunks[current], line);
    emit_bytes_to_string(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    set_global(&mut chunks[current], BODY_CACHE_GLOBAL, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// The query string parsed into a map. Stack: [] → [map].
///
/// Backs PHP `$_GET`, WSGI's `parse_qs(environ["QUERY_STRING"])`, Rack's
/// `GET` and ASP.NET `Request.Query`. Memoised so every language sees the
/// same object — a request is one request even when PHP calls into C#.
pub fn emit_query_params(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_memoised(
        chunks,
        current,
        QUERY_PARAMS_CACHE_GLOBAL,
        line,
        |chunks, current, line| {
            emit_query_string(chunks, current, line);
            super::http_form::emit_parse_urlencoded(chunks, current, line);
        },
    );
}

const QUERY_PARAMS_CACHE_GLOBAL: &str = "__vybe_request_query_params";

/// Run `build` once per request and keep the result in `cache_global`.
///
/// Shared identity is the point, not speed: `documentation/httpserver.md` §4a
/// requires that a request-derived surface be ONE object across languages, so
/// PHP writing `$_GET["x"]` and Python reading its query map see the same
/// thing in one process.
pub fn emit_memoised(
    chunks: &mut [Chunk],
    current: usize,
    cache_global: &str,
    line: u32,
    build: impl FnOnce(&mut [Chunk], usize, u32),
) {
    let out = chunks[current].alloc_scratch(1);

    push_global(&mut chunks[current], cache_global, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    push_global(&mut chunks[current], cache_global, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_else(line);
    build(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    set_global(&mut chunks[current], cache_global, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Query, body and cookie values merged into one map. Stack: [] → [map].
///
/// PHP's `$_REQUEST`. Later sources win, matching the order the values are
/// layered here; anything reading this wants "the request's parameters"
/// without caring how they arrived.
pub fn emit_request_params(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_memoised(
        chunks,
        current,
        REQUEST_PARAMS_CACHE_GLOBAL,
        line,
        |chunks, current, line| {
            let out = chunks[current].alloc_scratch(1);
            super::collections::emit_map_new(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

            emit_query_params(chunks, current, line);
            emit_merge_into(chunks, current, out, line);
            super::http_form::emit_body_fields(chunks, current, line);
            emit_merge_into(chunks, current, out, line);
            super::http_cookie::emit_request_cookies(chunks, current, line);
            emit_merge_into(chunks, current, out, line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        },
    );
}

const REQUEST_PARAMS_CACHE_GLOBAL: &str = "__vybe_request_params";

/// Copy every entry of the map on the stack into `target`. Stack: [map] → [].
fn emit_merge_into(chunks: &mut [Chunk], current: usize, target: u16, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, source, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    let map_keys = chunks[current].add_import("ecma:map", "keys");
    chunks[current].emit_call(map_keys, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
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

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    super::collections::emit_get(chunks, current, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_end(line);
}
