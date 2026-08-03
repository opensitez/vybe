//! Cookie parsing and serialization — the piece no spec surface provides.
//!
//! `wasi:http` carries headers, not cookies; Node has no cookie API either
//! (`cookie`/`express` are userland). RFC 6265 parsing therefore has to live
//! somewhere, and per `documentation/httpserver.md` §4a that somewhere is a
//! PRIMITIVE — emitted once, over the spec surfaces — rather than a host
//! function or, as it was, Rust inside the vybex server that only PHP could
//! reach.
//!
//! One implementation serves PHP `$_COOKIE`, Python `http.cookies`, Rack's
//! `cookies`, ASP.NET `Request.Cookies` and JS `req.cookies`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Parse a `Cookie:` header value into a map. Stack: [string] → [map].
///
/// RFC 6265 §4.2.1: `cookie-pair *( ";" SP cookie-pair )`. A pair with no `=`
/// is skipped rather than stored under an empty name, and surrounding spaces
/// are trimmed because the separator is `"; "`.
pub fn emit_parse_cookie_header(chunks: &mut [Chunk], current: usize, line: u32) {
    let header = chunks[current].alloc_scratch(1);
    let parts = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let eq = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, header, line);
    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    // A missing Cookie header is an empty map, not an error.
    chunks[current].emit_op_u16(Op::LOCAL_GET, header, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, header, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, header, line);
    chunks[current].emit_string_const(";", line);
    let split = chunks[current].add_import("ecma:string", "split");
    chunks[current].emit_call(split, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parts, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, parts, line);
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

    chunks[current].emit_op_u16(Op::LOCAL_GET, parts, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    let trim = chunks[current].add_import("ecma:string", "trim");
    chunks[current].emit_call(trim, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    chunks[current].emit_string_const("=", line);
    let index_of = chunks[current].add_import("ecma:string", "indexOf");
    chunks[current].emit_call(index_of, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, eq, line);

    // Only store when there is an `=`; a bare token is not a cookie-pair.
    chunks[current].emit_op_u16(Op::LOCAL_GET, eq, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    // name = pair[0..eq]
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, eq, line);
    let substring = chunks[current].add_import("wasm:js-string", "substring");
    chunks[current].emit_call(substring, 3, line);
    // value = pair[eq+1..len]
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, eq, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    let len = chunks[current].add_import("wasm:js-string", "length");
    chunks[current].emit_call(len, 1, line);
    let substring2 = chunks[current].add_import("wasm:js-string", "substring");
    chunks[current].emit_call(substring2, 3, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);

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

/// This request's cookies as a map. Stack: [] → [map].
///
/// Reads the `Cookie` header off the `wasi:http` request, so it works for any
/// language and needs no per-language request context.
pub fn emit_request_cookies(chunks: &mut [Chunk], current: usize, line: u32) {
    super::http_request_env::emit_memoised(
        chunks,
        current,
        COOKIE_CACHE_GLOBAL,
        line,
        |chunks, current, line| {
            super::http_request_env::emit_header(chunks, current, "cookie", line);
            emit_parse_cookie_header(chunks, current, line);
        },
    );
}

/// Parsed once per request: PHP's `$_COOKIE`, Python's `http.cookies` view and
/// ASP.NET's `Request.Cookies` have to be the same object, or a cookie written
/// through one is invisible to the next.
const COOKIE_CACHE_GLOBAL: &str = "__vybe_request_cookies";

/// Build a `Set-Cookie` header VALUE. Stack: [name, value, attrs?] → [string].
///
/// RFC 6265 §4.1.1 serialization — the mirror of `emit_parse_cookie_header`,
/// and the reason cookie writing no longer needs a host function. `attrs` is an
/// optional map with any of `expires` (unix seconds), `path`, `domain`,
/// `samesite`, `secure`, `httponly`; that is the shape PHP 7.3+ `setcookie()`
/// takes, and also what Rack, ASP.NET and `express` accept.
///
/// The value is written VERBATIM. Whether it gets URL-encoded is the calling
/// language's business — PHP's `setcookie()` encodes and `setrawcookie()` does
/// not — and baking either choice in here would make the primitive PHP's.
pub fn emit_serialize(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let attrs = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let expires = chunks[current].alloc_scratch(1);
    // One slot per attribute, allocated UP FRONT. Allocating mid-emission
    // interleaves with the slots nested emit helpers take for themselves.
    let path_slot = chunks[current].alloc_scratch(1);
    let domain_slot = chunks[current].alloc_scratch(1);
    let samesite_slot = chunks[current].alloc_scratch(1);
    let secure_slot = chunks[current].alloc_scratch(1);
    let httponly_slot = chunks[current].alloc_scratch(1);

    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, attrs, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, attrs, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    chunks[current].emit_string_const("=", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    super::strings::emit_concat(&mut chunks[current], 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, attrs, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    // `Expires` is an HTTP-date, not the unix timestamp callers pass in.
    // `Date.prototype.toUTCString` produces exactly that format (RFC 9110
    // §5.6.7 IMF-fixdate), so no date formatting is hand-rolled here.
    emit_attr(chunks, current, attrs, "expires", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, expires, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, expires, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, expires, line);
    let to_number = chunks[current].add_import("ecma:value", "toNumber");
    chunks[current].emit_call(to_number, 1, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const("; Expires=", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, expires, line);
    let to_number = chunks[current].add_import("ecma:value", "toNumber");
    chunks[current].emit_call(to_number, 1, line);
    chunks[current].emit_f64_const(1000.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    let date_new = chunks[current].add_import("ecma:date", "new");
    chunks[current].emit_call(date_new, 1, line);
    let to_utc = chunks[current].add_import("ecma:date", "toUTCString");
    chunks[current].emit_call(to_utc, 1, line);
    super::strings::emit_concat(&mut chunks[current], 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    for (key, label, slot) in [
        ("path", "; Path=", path_slot),
        ("domain", "; Domain=", domain_slot),
        ("samesite", "; SameSite=", samesite_slot),
    ] {
        emit_attr(chunks, current, attrs, key, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_string_const(label, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        super::strings::emit_concat(&mut chunks[current], 3, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
        chunks[current].emit_end(line);
    }

    // `Secure` and `HttpOnly` are valueless flags — present or absent.
    for (key, label, slot) in [
        ("secure", "; Secure", secure_slot),
        ("httponly", "; HttpOnly", httponly_slot),
    ] {
        emit_attr(chunks, current, attrs, key, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        super::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_string_const(label, line);
        super::strings::emit_concat(&mut chunks[current], 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }

    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Read one attribute off the attrs map. Stack: [] → [value|null].
fn emit_attr(chunks: &mut [Chunk], current: usize, attrs: u16, key: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, attrs, line);
    chunks[current].emit_string_const(key, line);
    super::collections::emit_get(chunks, current, line);
}
