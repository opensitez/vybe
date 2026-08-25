//! Python `zlib` / `gzip` stdlib adapters.
//!
//! The compressors themselves are the existing `node:zlib` host surface
//! (flate2), bound straight from the profile — `zlib.compress` is a
//! `host:node:zlib:deflateSync` row, not code. What lives here is the part
//! `node:zlib` does NOT export: Adler-32, the checksum RFC 1950 puts in a
//! zlib header, which `zlib.adler32` exposes on its own.
//!
//! Adler-32 (RFC 1950 §9): two 16-bit sums mod 65521, packed `(b << 16) | a`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::{collections, loops};

/// Largest prime below 2^16 — the Adler-32 modulus.
const ADLER_MOD: i32 = 65521;

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

/// `zlib.adler32(data[, value])` → unsigned 32-bit checksum.
pub fn emit_adler32(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc.max(1), line);
    let bytes = base;
    let a = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);

    // A continuation `value` splits back into the two running sums; without
    // one the checksum starts at 1 (RFC 1950: `a` seeds to 1, `b` to 0).
    if argc >= 2 {
        lget(&mut chunks[current], base + 1, line);
        chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
        chunks[current].emit_i32_const(0xFFFF, line);
        chunks[current].emit_op(Op::I32_AND, line);
        lset(&mut chunks[current], a, line);
        lget(&mut chunks[current], base + 1, line);
        chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
        chunks[current].emit_i32_const(16, line);
        chunks[current].emit_op(Op::I32_SHR_U, line);
        chunks[current].emit_i32_const(0xFFFF, line);
        chunks[current].emit_op(Op::I32_AND, line);
        lset(&mut chunks[current], b, line);
    } else {
        chunks[current].emit_i32_const(1, line);
        lset(&mut chunks[current], a, line);
        chunks[current].emit_i32_const(0, line);
        lset(&mut chunks[current], b, line);
    }

    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);
    lget(&mut chunks[current], bytes, line);
    collections::emit_len(chunks, current, line);
    lset(&mut chunks[current], len, line);

    let loop_id = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    // a = (a + byte) % 65521
    lget(&mut chunks[current], a, line);
    lget(&mut chunks[current], bytes, line);
    lget(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    let idx = chunks[current].add_import("wasm:js-number", "toI32");
    chunks[current].emit_call(idx, 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_i32_const(ADLER_MOD, line);
    chunks[current].emit_op(Op::I32_REM_U, line);
    lset(&mut chunks[current], a, line);

    // b = (b + a) % 65521
    lget(&mut chunks[current], b, line);
    lget(&mut chunks[current], a, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_i32_const(ADLER_MOD, line);
    chunks[current].emit_op(Op::I32_REM_U, line);
    lset(&mut chunks[current], b, line);

    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    loops::emit_loop_end(chunks, current, loop_id, line);

    lget(&mut chunks[current], b, line);
    chunks[current].emit_i32_const(16, line);
    chunks[current].emit_op(Op::I32_SHL, line);
    lget(&mut chunks[current], a, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
}

/// `node:zlib` hands back a plain `ObjectKind::Array` of byte values, but
/// Python's compressors return `bytes` — the same `Uint8Array` a `b'…'`
/// literal builds. Stack: `[array]` → `[bytes]`.
fn to_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("ecma:uint8array", "new");
    chunks[current].emit_call(idx, 1, line);
}

/// `zlib.compress` / `gzip.compress`. Python passes the level as a bare int;
/// `node:zlib` reads it off an options object, so the adapter is what bridges
/// the two shapes. `default_level` is the level Python uses when the call
/// omits one — `gzip.compress` defaults to 9, `zlib.compress` to the
/// library default (-1).
fn emit_compress_with(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    func: &str,
    default_level: i32,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(2);
    let data = base;
    let level = base + 1;
    if argc >= 2 {
        for offset in (0..2u16).rev() {
            lset(&mut chunks[current], base + offset, line);
        }
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    } else {
        lset(&mut chunks[current], data, line);
        chunks[current].emit_i32_const(default_level, line);
        lset(&mut chunks[current], level, line);
    }

    lget(&mut chunks[current], data, line);
    // `{ level: <n> }` — the options object node's zlib reads.
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    lget(&mut chunks[current], level, line);
    let key = chunks[current].add_constant(vybe_runtime::Value::String(
        std::sync::Arc::from("level"),
    ));
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    let idx = chunks[current].add_import("node:zlib", func);
    chunks[current].emit_call(idx, 2, line);
    to_bytes(chunks, current, line);
}

/// `zlib.decompress` / `gzip.decompress`. Python's trailing `wbits` /
/// `bufsize` arguments only size the internal buffer, which the host
/// decompressor grows on its own, so they are dropped.
fn emit_decompress_with(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    func: &str,
    line: u32,
) {
    let data = chunks[current].alloc_scratch(1);
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    lset(&mut chunks[current], data, line);
    lget(&mut chunks[current], data, line);
    let idx = chunks[current].add_import("node:zlib", func);
    chunks[current].emit_call(idx, 1, line);
    to_bytes(chunks, current, line);
}

/// `zlib.compress(data[, level])` — RFC 1950 (zlib wrapper).
pub fn emit_zlib_compress(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_compress_with(chunks, current, argc, "deflateSync", -1, line);
}

/// `zlib.decompress(data)`.
pub fn emit_zlib_decompress(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_decompress_with(chunks, current, argc, "inflateSync", line);
}

/// `gzip.compress(data[, compresslevel])` — RFC 1952 (gzip wrapper).
pub fn emit_gzip_compress(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_compress_with(chunks, current, argc, "gzipSync", 9, line);
}

/// `gzip.decompress(data)`.
pub fn emit_gzip_decompress(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_decompress_with(chunks, current, argc, "gunzipSync", line);
}
