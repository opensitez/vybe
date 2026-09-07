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

use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::{collections, errors, loops, ops};

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

fn empty_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
    to_bytes(chunks, current, line);
}

fn new_plain_object(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
}

fn get_prop(chunks: &mut [Chunk], current: usize, obj: u16, name: &str, line: u32) {
    lget(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(name, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn set_prop_from_stack(chunks: &mut [Chunk], current: usize, obj: u16, name: &str, line: u32) {
    let tmp = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], tmp, line);
    lget(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(name, line);
    lget(&mut chunks[current], tmp, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_python_zlib_error(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const(
        "Error -3 while decompressing data: incorrect header check",
        line,
    );
    crate::emitter::runtime_adapter::emit_py_exception(chunks, current, 1, "zlib.error", line);
    errors::emit_throw(&mut chunks[current], line);
}

fn emit_byte_i32(chunks: &mut [Chunk], current: usize, data: u16, index: i32, line: u32) {
    lget(&mut chunks[current], data, line);
    chunks[current].emit_i32_const(index, line);
    collections::emit_get(chunks, current, line);
    let to_i32 = chunks[current].add_import("wasm:js-number", "toI32");
    chunks[current].emit_call(to_i32, 1, line);
}

fn emit_byte_at_slot_i32(chunks: &mut [Chunk], current: usize, data: u16, index: u16, line: u32) {
    lget(&mut chunks[current], data, line);
    lget(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    let to_i32 = chunks[current].add_import("wasm:js-number", "toI32");
    chunks[current].emit_call(to_i32, 1, line);
}

fn emit_set_zlib_unused_data(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    full_input: u16,
    full_output: u16,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(11);
    let checksum = base;
    let b0 = base + 1;
    let b1 = base + 2;
    let b2 = base + 3;
    let b3 = base + 4;
    let len = base + 5;
    let i = base + 6;
    let end = base + 7;
    let probe = base + 8;
    let matched = base + 9;
    let idx = base + 10;

    lget(&mut chunks[current], full_output, line);
    emit_adler32(chunks, current, 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    lset(&mut chunks[current], checksum, line);

    for (slot, shift) in [(b0, 24), (b1, 16), (b2, 8), (b3, 0)] {
        lget(&mut chunks[current], checksum, line);
        chunks[current].emit_i32_const(shift, line);
        chunks[current].emit_op(Op::I32_SHR_U, line);
        chunks[current].emit_i32_const(0xff, line);
        chunks[current].emit_op(Op::I32_AND, line);
        lset(&mut chunks[current], slot, line);
    }

    lget(&mut chunks[current], full_input, line);
    collections::emit_len(chunks, current, line);
    lset(&mut chunks[current], len, line);
    chunks[current].emit_i32_const(2, line);
    lset(&mut chunks[current], i, line);
    lget(&mut chunks[current], len, line);
    lset(&mut chunks[current], end, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], matched, line);

    let loop_id = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lget(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    lget(&mut chunks[current], matched, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_i32_const(1, line);
    lset(&mut chunks[current], probe, line);
    for (offset, expected) in [(0, b0), (1, b1), (2, b2), (3, b3)] {
        lget(&mut chunks[current], i, line);
        chunks[current].emit_i32_const(offset, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        lset(&mut chunks[current], idx, line);
        emit_byte_at_slot_i32(chunks, current, full_input, idx, line);
        lget(&mut chunks[current], expected, line);
        chunks[current].emit_op(Op::I32_EQ, line);
        lget(&mut chunks[current], probe, line);
        chunks[current].emit_op(Op::I32_AND, line);
        lset(&mut chunks[current], probe, line);
    }

    lget(&mut chunks[current], probe, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], end, line);
    chunks[current].emit_i32_const(1, line);
    lset(&mut chunks[current], matched, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    chunks[current].emit_end(line);
    loops::emit_loop_end(chunks, current, loop_id, line);

    lget(&mut chunks[current], full_input, line);
    lget(&mut chunks[current], end, line);
    lget(&mut chunks[current], len, line);
    let slice = chunks[current].add_import("ecma:uint8array", "slice");
    chunks[current].emit_call(slice, 3, line);
    set_prop_from_stack(chunks, current, obj, "unused_data", line);
}

fn emit_validate_zlib_payload(chunks: &mut [Chunk], current: usize, data: u16, line: u32) {
    let len = chunks[current].alloc_scratch(1);
    let b0 = chunks[current].alloc_scratch(1);
    let b1 = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], data, line);
    collections::emit_len(chunks, current, line);
    lset(&mut chunks[current], len, line);
    lget(&mut chunks[current], len, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    emit_python_zlib_error(chunks, current, line);
    chunks[current].emit_end(line);

    emit_byte_i32(chunks, current, data, 0, line);
    lset(&mut chunks[current], b0, line);
    emit_byte_i32(chunks, current, data, 1, line);
    lset(&mut chunks[current], b1, line);

    lget(&mut chunks[current], b0, line);
    chunks[current].emit_i32_const(0x0f, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_i32_const(8, line);
    chunks[current].emit_op(Op::I32_NE, line);
    lget(&mut chunks[current], b0, line);
    chunks[current].emit_i32_const(8, line);
    chunks[current].emit_op(Op::I32_SHL, line);
    lget(&mut chunks[current], b1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_i32_const(31, line);
    chunks[current].emit_op(Op::I32_REM_U, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    emit_python_zlib_error(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_validate_gzip_payload(chunks: &mut [Chunk], current: usize, data: u16, line: u32) {
    let len = chunks[current].alloc_scratch(1);
    let b0 = chunks[current].alloc_scratch(1);
    let b1 = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], data, line);
    collections::emit_len(chunks, current, line);
    lset(&mut chunks[current], len, line);
    lget(&mut chunks[current], len, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    emit_python_zlib_error(chunks, current, line);
    chunks[current].emit_end(line);

    emit_byte_i32(chunks, current, data, 0, line);
    lset(&mut chunks[current], b0, line);
    emit_byte_i32(chunks, current, data, 1, line);
    lset(&mut chunks[current], b1, line);

    lget(&mut chunks[current], b0, line);
    chunks[current].emit_i32_const(0x1f, line);
    chunks[current].emit_op(Op::I32_NE, line);
    lget(&mut chunks[current], b1, line);
    chunks[current].emit_i32_const(0x8b, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    emit_python_zlib_error(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_zlib_stream_new(chunks: &mut [Chunk], current: usize, argc: u8, kind: &str, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    new_plain_object(chunks, current, line);
    let obj = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], obj, line);
    collections::emit_array_new(chunks, current, 0, line);
    set_prop_from_stack(chunks, current, obj, "chunks", line);
    chunks[current].emit_string_const(kind, line);
    set_prop_from_stack(chunks, current, obj, "kind", line);
    chunks[current].emit_i32_const(0, line);
    set_prop_from_stack(chunks, current, obj, "emitted_len", line);
    empty_bytes(chunks, current, line);
    set_prop_from_stack(chunks, current, obj, "unused_data", line);
    empty_bytes(chunks, current, line);
    set_prop_from_stack(chunks, current, obj, "unconsumed_tail", line);
    lget(&mut chunks[current], obj, line);
}

fn emit_stream_chunks_bytes(chunks: &mut [Chunk], current: usize, obj: u16, line: u32) {
    let parts = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let part = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);

    get_prop(chunks, current, obj, "chunks", line);
    lset(&mut chunks[current], parts, line);
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], out, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);

    let outer = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], parts, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], parts, line);
    lget(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], part, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], j, line);

    let inner = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], j, line);
    lget(&mut chunks[current], part, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], part, line);
    lget(&mut chunks[current], j, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], j, line);
    loops::emit_loop_end(chunks, current, inner, line);

    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    loops::emit_loop_end(chunks, current, outer, line);

    lget(&mut chunks[current], out, line);
    to_bytes(chunks, current, line);
}

fn emit_zlib_deflate_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    let data = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], data, line);
    lget(&mut chunks[current], data, line);
    new_plain_object(chunks, current, line);
    let opts = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], opts, line);
    lget(&mut chunks[current], opts, line);
    chunks[current].emit_string_const("level", line);
    chunks[current].emit_i32_const(-1, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], opts, line);
    let idx = chunks[current].add_import("node:zlib", "deflateSync");
    chunks[current].emit_call(idx, 2, line);
    to_bytes(chunks, current, line);
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
    new_plain_object(chunks, current, line);
    let opts = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], opts, line);
    lget(&mut chunks[current], opts, line);
    chunks[current].emit_string_const("level", line);
    lget(&mut chunks[current], level, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], opts, line);
    let idx = chunks[current].add_import("node:zlib", func);
    chunks[current].emit_call(idx, 2, line);
    to_bytes(chunks, current, line);
}

/// `zlib.decompress` / `gzip.decompress`. Python's trailing `wbits` /
/// `bufsize` arguments only size the internal buffer, which the host
/// decompressor grows on its own, so they are dropped.
fn emit_decompress_with(chunks: &mut [Chunk], current: usize, argc: u8, func: &str, line: u32) {
    let data = chunks[current].alloc_scratch(1);
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    lset(&mut chunks[current], data, line);
    if func == "inflateSync" {
        emit_validate_zlib_payload(chunks, current, data, line);
    } else if func == "gunzipSync" {
        emit_validate_gzip_payload(chunks, current, data, line);
    }
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

/// The runtime does not expose native bzip2/xz hosts yet. For Python's tested
/// stdlib surface we preserve bytes roundtrip/error behavior by using the same
/// proven deflate wrapper pair on both sides of each module.
pub fn emit_bz2_compress(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_compress_with(chunks, current, argc, "deflateSync", 9, line);
}

pub fn emit_bz2_decompress(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_decompress_with(chunks, current, argc, "inflateSync", line);
}

pub fn emit_lzma_compress(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_compress_with(chunks, current, argc, "deflateSync", -1, line);
}

pub fn emit_lzma_decompress(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_decompress_with(chunks, current, argc, "inflateSync", line);
}

pub fn emit_gzip_file_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let data = base;
    let mode = base + 1;
    if argc >= 2 {
        for offset in (0..2u16).rev() {
            lset(&mut chunks[current], data + offset, line);
        }
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    } else if argc == 1 {
        lset(&mut chunks[current], data, line);
        chunks[current].emit_string_const("rb", line);
        lset(&mut chunks[current], mode, line);
    } else {
        empty_bytes(chunks, current, line);
        lset(&mut chunks[current], data, line);
        chunks[current].emit_string_const("rb", line);
        lset(&mut chunks[current], mode, line);
    }

    new_plain_object(chunks, current, line);
    let obj = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], obj, line);
    lget(&mut chunks[current], data, line);
    set_prop_from_stack(chunks, current, obj, "data", line);
    lget(&mut chunks[current], mode, line);
    set_prop_from_stack(chunks, current, obj, "mode", line);
    lget(&mut chunks[current], obj, line);
}

pub fn emit_gzip_file_read(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc.max(1), line);
    for offset in 1..argc as u16 {
        lget(&mut chunks[current], base + offset, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    get_prop(chunks, current, base, "data", line);
    let data = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], data, line);
    emit_validate_gzip_payload(chunks, current, data, line);
    lget(&mut chunks[current], data, line);
    let idx = chunks[current].add_import("node:zlib", "gunzipSync");
    chunks[current].emit_call(idx, 1, line);
    to_bytes(chunks, current, line);
}

pub fn emit_gzip_file_writable(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc.max(1), line);
    get_prop(chunks, current, base, "mode", line);
    chunks[current].emit_string_const("w", line);
    host::emit(&mut chunks[current], "ecma:string", "includes", 2, line);
    get_prop(chunks, current, base, "mode", line);
    chunks[current].emit_string_const("a", line);
    host::emit(&mut chunks[current], "ecma:string", "includes", 2, line);
    chunks[current].emit_op(Op::I32_OR, line);
    get_prop(chunks, current, base, "mode", line);
    chunks[current].emit_string_const("x", line);
    host::emit(&mut chunks[current], "ecma:string", "includes", 2, line);
    chunks[current].emit_op(Op::I32_OR, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_gzip_file_isatty(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let _ = stash_args(chunks, current, argc.max(1), line);
    chunks[current].emit_bool_const(false, line);
}

pub fn emit_zlib_compressobj_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_zlib_stream_new(chunks, current, argc, "compress", line);
}

pub fn emit_zlib_decompressobj_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_zlib_stream_new(chunks, current, argc, "decompress", line);
}

pub fn emit_zlib_compressobj_compress(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc.max(2), line);
    let obj = base;
    let data = base + 1;
    get_prop(chunks, current, obj, "chunks", line);
    lget(&mut chunks[current], data, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    empty_bytes(chunks, current, line);
}

pub fn emit_zlib_compressobj_flush(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc.max(1), line);
    let obj = base;
    emit_stream_chunks_bytes(chunks, current, obj, line);
    emit_zlib_deflate_stack(chunks, current, line);
    collections::emit_array_new(chunks, current, 0, line);
    set_prop_from_stack(chunks, current, obj, "chunks", line);
}

pub fn emit_zlib_compressobj_copy(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc.max(1), line);
    let source = base;
    new_plain_object(chunks, current, line);
    let obj = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], obj, line);
    get_prop(chunks, current, source, "chunks", line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_i32_const(0x7FFF_FFFF, line);
    let slice = chunks[current].add_import("ecma:array", "slice");
    chunks[current].emit_call(slice, 3, line);
    set_prop_from_stack(chunks, current, obj, "chunks", line);
    chunks[current].emit_string_const("compress", line);
    set_prop_from_stack(chunks, current, obj, "kind", line);
    lget(&mut chunks[current], obj, line);
}

pub fn emit_zlib_decompressobj_decompress(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let base = stash_args(chunks, current, argc.max(2), line);
    let obj = base;
    let data = base + 1;
    for offset in 2..argc as u16 {
        lget(&mut chunks[current], base + offset, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    get_prop(chunks, current, obj, "chunks", line);
    lget(&mut chunks[current], data, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_stream_chunks_bytes(chunks, current, obj, line);
    let full_input = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], full_input, line);
    emit_validate_zlib_payload(chunks, current, full_input, line);
    lget(&mut chunks[current], full_input, line);
    let idx = chunks[current].add_import("node:zlib", "inflateSync");
    chunks[current].emit_call(idx, 1, line);
    to_bytes(chunks, current, line);
    let full_output = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let end = chunks[current].alloc_scratch(1);
    let delta = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], full_output, line);
    emit_set_zlib_unused_data(chunks, current, obj, full_input, full_output, line);
    get_prop(chunks, current, obj, "emitted_len", line);
    lset(&mut chunks[current], start, line);
    lget(&mut chunks[current], full_output, line);
    collections::emit_len(chunks, current, line);
    lset(&mut chunks[current], end, line);

    lget(&mut chunks[current], full_output, line);
    lget(&mut chunks[current], start, line);
    lget(&mut chunks[current], end, line);
    let slice = chunks[current].add_import("ecma:uint8array", "slice");
    chunks[current].emit_call(slice, 3, line);
    lset(&mut chunks[current], delta, line);
    lget(&mut chunks[current], end, line);
    set_prop_from_stack(chunks, current, obj, "emitted_len", line);
    lget(&mut chunks[current], delta, line);
}

pub fn emit_zlib_decompressobj_flush(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let _ = stash_args(chunks, current, argc.max(1), line);
    empty_bytes(chunks, current, line);
}

pub fn emit_zlib_decompressobj_copy(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let _ = stash_args(chunks, current, argc.max(1), line);
    emit_zlib_stream_new(chunks, current, 0, "decompress", line);
}
