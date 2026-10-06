// Frozen URLencoded form emitter before cached separator reuse.
//! Request body parsing — `application/x-www-form-urlencoded` and
//! `multipart/form-data`.
//!
//! Neither is in any spec host surface. `wasi:http` carries the body as bytes;
//! `node:querystring` covers urlencoded only, and **multipart exists nowhere**.
//! Per `documentation/httpserver.md` §4a this is therefore a PRIMITIVE, and
//! deliberately ONE of them: multipart parsing has a real CVE history, so the
//! project ships a single implementation rather than one per language.
//!
//! One implementation serves PHP `$_POST`/`$_FILES`, Python
//! `cgi.FieldStorage`, Rack's multipart, ASP.NET `IFormFile` and JS
//! `multipart/form-data`.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// Parse `a=1&b=2` into a map. Stack: [string] → [map].
///
/// Percent-decoding and `+`-as-space are the urlencoded rules (RFC 1866
/// §8.2.1), applied through `ecma:string.decodeURIComponent` rather than
/// hand-rolled.
pub fn emit_parse_urlencoded(chunks: &mut [Chunk], current: usize, line: u32) {
    let src = chunks[current].alloc_scratch(1);
    let parts = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let eq = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);
    vybe_compiler::primitives::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    // A null or empty body is an empty map, not an error.
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_string_const("&", line);
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
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    chunks[current].emit_string_const("=", line);
    let index_of = chunks[current].add_import("ecma:string", "indexOf");
    chunks[current].emit_call(index_of, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, eq, line);

    // A field with no `=` is skipped; an empty segment (`a=1&&b=2`) too.
    chunks[current].emit_op_u16(Op::LOCAL_GET, eq, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    emit_decode_slice(chunks, current, pair, 0, Some(eq), line);
    emit_decode_slice(chunks, current, pair, 1, None, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
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

/// Push `pair`'s name (`part` 0, up to `eq`) or value (`part` 1, after `eq`),
/// `+`-decoded and percent-decoded. Stack: [] → [string].
fn emit_decode_slice(
    chunks: &mut [Chunk],
    current: usize,
    pair: u16,
    part: u8,
    eq: Option<u16>,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    if part == 0 {
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, eq.expect("name slice needs eq"), line);
    } else {
        // The `=` index is recomputed here so the caller need not keep it live
        // across the name push.
        chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
        chunks[current].emit_string_const("=", line);
        let index_of = chunks[current].add_import("ecma:string", "indexOf");
        chunks[current].emit_call(index_of, 2, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
        let len = chunks[current].add_import("wasm:js-string", "length");
        chunks[current].emit_call(len, 1, line);
    }
    let substring = chunks[current].add_import("wasm:js-string", "substring");
    chunks[current].emit_call(substring, 3, line);

    // `+` means space before percent-decoding (RFC 1866 §8.2.1).
    chunks[current].emit_string_const("+", line);
    chunks[current].emit_string_const(" ", line);
    let replace_all = chunks[current].add_import("ecma:string", "replaceAll");
    chunks[current].emit_call(replace_all, 3, line);

    let decode = chunks[current].add_import("ecma:string", "decodeURIComponent");
    chunks[current].emit_call(decode, 1, line);
}
