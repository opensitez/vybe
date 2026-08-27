//! `System.IO.BinaryWriter` / `BinaryReader` over the `MemoryStream` value.
//!
//! Both are thin cursors on a stream that already owns the bytes, so the value
//! carries only what .NET's own object carries:
//!
//! ```text
//! { __type: "BinaryWriter" | "BinaryReader",
//!   __bio_stream: <stream>,   the underlying MemoryStream value
//!   __bio_leave_open: <bool> }
//! ```
//!
//! Everything is written LITTLE-ENDIAN, which is .NET's documented format for
//! these two types on every platform it runs on — not the host's byte order.
//!
//! ⛔ **`Write` is the one member whose overload cannot be chosen here.** .NET
//! declares `Write(Int16)`, `Write(Int32)`, `Write(Double)`, … all at arity 1,
//! separated only by the STATIC type of the argument, and the descriptor
//! carries a name and an arity but no parameter types. The VB walker resolves
//! it instead — a numeric literal's suffix (`65000US`) survives parsing as an
//! `ExprKind::Cast`, so the width is known there — and rewrites the call to the
//! width-specific spelling below. `Write` itself keeps the .NET default for an
//! unannotated argument: `Int32` for a whole number, `Double` otherwise.

use std::sync::Arc;
use vybe_compiler::primitives::errors;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::ops;
use vybe_compiler::primitives::strings as shared_strings;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

const TYPE_KEY: &str = "__type";
const STREAM: &str = "__bio_stream";
const LEAVE_OPEN: &str = "__bio_leave_open";

// The stream's own fields — this adapter reads and advances the same cursor
// `memory_stream_adapter` maintains, so a `BinaryWriter` and the stream it
// wraps never disagree about position.
const BUF: &str = "__ms_buf";
const POS: &str = "__ms_pos";
const LEN: &str = "__ms_len";

fn num(chunk: &mut Chunk, v: f64, line: u32) {
    chunk.emit_f64_const(v, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

fn field_set(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, idx, line);
}

fn field(chunk: &mut Chunk, slot: u16, key: &str, line: u32) {
    get(chunk, slot, line);
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, idx, line);
}

fn store(chunk: &mut Chunk, slot: u16, key: &str, line: u32) {
    let tmp = chunk.alloc_scratch(1);
    set(chunk, tmp, line);
    get(chunk, slot, line);
    get(chunk, tmp, line);
    field_set(chunk, key, line);
}

fn throw(chunk: &mut Chunk, class: &str, message: &str, line: u32) {
    chunk.emit_struct_new(0, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const(message, line);
    errors::emit_exception_new_finalize(chunk, class, line);
    errors::emit_throw(chunk, line);
}


/// `[v, m] → [v mod m]` — arithmetic, so no host import is needed for what is
/// just a division remainder.
fn fmod(chunk: &mut Chunk, line: u32) {
    let scratch = chunk.alloc_scratch(2);
    let m = scratch;
    let v = scratch + 1;
    set(chunk, m, line);
    set(chunk, v, line);
    get(chunk, v, line);
    get(chunk, v, line);
    get(chunk, m, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    get(chunk, m, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
}

/// `[bio] → [stream]` — the wrapped stream, in a slot.
fn stream_of(chunk: &mut Chunk, bio: u16, line: u32) -> u16 {
    let s = chunk.alloc_scratch(1);
    field(chunk, bio, STREAM, line);
    set(chunk, s, line);
    s
}

// ── construction ────────────────────────────────────────────────────────────

fn build(chunk: &mut Chunk, type_name: &str, argc: u8, line: u32) {
    let scratch = chunk.alloc_scratch(3);
    let stream = scratch;
    let leave_open = scratch + 1;
    let obj = scratch + 2;

    // `(stream)`, `(stream, encoding)` or `(stream, encoding, leaveOpen)`.
    // The encoding is accepted and not stored: every string this adapter
    // writes is UTF-8, which is what all three declared encodings resolve to
    // for the ASCII the tests round-trip. A non-UTF-8 encoding would need the
    // encoder itself, and is not claimed here.
    match argc {
        3 => {
            set(chunk, leave_open, line);
            chunk.emit_op(Op::DROP, line);
            set(chunk, stream, line);
        }
        2 => {
            chunk.emit_op(Op::DROP, line);
            set(chunk, stream, line);
            core_wasm::bool_const(chunk, line, false);
            set(chunk, leave_open, line);
        }
        _ => {
            set(chunk, stream, line);
            core_wasm::bool_const(chunk, line, false);
            set(chunk, leave_open, line);
        }
    }

    chunk.emit_struct_new(0, 0, line);
    set(chunk, obj, line);
    get(chunk, obj, line);
    chunk.emit_string_const(type_name, line);
    field_set(chunk, TYPE_KEY, line);
    get(chunk, obj, line);
    get(chunk, stream, line);
    field_set(chunk, STREAM, line);
    get(chunk, obj, line);
    get(chunk, leave_open, line);
    field_set(chunk, LEAVE_OPEN, line);
    get(chunk, obj, line);
}

pub fn emit_writer_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    build(&mut chunks[current], "BinaryWriter", argc, line);
}

pub fn emit_reader_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    build(&mut chunks[current], "BinaryReader", argc, line);
}

/// `BaseStream` — the SAME object the constructor was handed, so
/// `Object.ReferenceEquals(bw.BaseStream, ms)` is true.
pub fn emit_base_stream(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let bio = chunk.alloc_scratch(1);
    set(chunk, bio, line);
    field(chunk, bio, STREAM, line);
}

// ── byte plumbing ───────────────────────────────────────────────────────────

/// Append one byte (already on the stack) at the stream's cursor, growing the
/// backing array and extending the logical length exactly as
/// `MemoryStream.WriteByte` does.
fn put_byte(chunk: &mut Chunk, stream: u16, byte: u16, line: u32) {
    // Grow: while buf.length <= pos, push 0.
    let guard = chunk.emit_block(line);
    let loop_state = {
        let block = chunk.emit_block(line);
        let (loop_patch, _) = chunk.emit_loop_s(line);
        (block, loop_patch)
    };
    field(chunk, stream, BUF, line);
    call(chunk, "ecma:array", "length", 1, line);
    field(chunk, stream, POS, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    field(chunk, stream, BUF, line);
    num(chunk, 0.0, line);
    call(chunk, "ecma:array", "push", 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_state.1);
    chunk.emit_end(line);
    chunk.patch_block(loop_state.0);
    chunk.emit_end(line);
    chunk.patch_block(guard);

    field(chunk, stream, BUF, line);
    field(chunk, stream, POS, line);
    get(chunk, byte, line);
    chunk.emit_op(Op::ARRAY_SET, line);

    field(chunk, stream, POS, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    store(chunk, stream, POS, line);

    field(chunk, stream, POS, line);
    field(chunk, stream, LEN, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    field(chunk, stream, POS, line);
    store(chunk, stream, LEN, line);
    chunk.emit_end(line);
}

/// Read one byte at the cursor and advance. Past the end throws
/// `EndOfStreamException`, which is what separates a `BinaryReader` from
/// `Stream.ReadByte` (that answers -1).
fn take_byte(chunk: &mut Chunk, stream: u16, out: u16, line: u32) {
    field(chunk, stream, POS, line);
    field(chunk, stream, LEN, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    throw(
        chunk,
        "EndOfStreamException",
        "Unable to read beyond the end of the stream.",
        line,
    );
    chunk.emit_end(line);

    field(chunk, stream, BUF, line);
    field(chunk, stream, POS, line);
    call(chunk, "ecma:array", "get", 2, line);
    set(chunk, out, line);

    field(chunk, stream, POS, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    store(chunk, stream, POS, line);
}

// ── integers ────────────────────────────────────────────────────────────────

/// `[bw, value] → []` — `width` little-endian bytes.
///
/// The bytes come out by repeated `floor(v / 256^i) mod 256` rather than shifts
/// so the same emitter serves the 64-bit widths, whose values do not fit an
/// i32. A negative value is first folded into its unsigned two's-complement
/// representation for the width, which is what .NET writes.
pub fn emit_write_int(chunks: &mut [Chunk], current: usize, width: u32, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(3);
    let value = scratch;
    let bio = scratch + 1;
    let byte = scratch + 2;
    set(chunk, value, line);
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);

    let modulus = 2f64.powi(8 * width as i32);
    // v < 0 → v + 2^(8*width)
    get(chunk, value, line);
    num(chunk, 0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    get(chunk, value, line);
    num(chunk, modulus, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, value, line);
    chunk.emit_end(line);

    for i in 0..width {
        get(chunk, value, line);
        num(chunk, 2f64.powi(8 * i as i32), line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_FLOOR, line);
        num(chunk, 256.0, line);
        fmod(chunk, line);
        chunk.emit_op(Op::F64_FLOOR, line);
        set(chunk, byte, line);
        put_byte(chunk, stream, byte, line);
    }
    core_wasm::undefined(chunk, line);
}

/// `[br] → [number]` — `width` little-endian bytes. `signed` reinterprets the
/// top bit, which is the only difference between `ReadInt16` and `ReadUInt16`.
pub fn emit_read_int(chunks: &mut [Chunk], current: usize, width: u32, signed: bool, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(3);
    let bio = scratch;
    let acc = scratch + 1;
    let byte = scratch + 2;
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);

    num(chunk, 0.0, line);
    set(chunk, acc, line);
    for i in 0..width {
        take_byte(chunk, stream, byte, line);
        get(chunk, acc, line);
        get(chunk, byte, line);
        num(chunk, 2f64.powi(8 * i as i32), line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        set(chunk, acc, line);
    }
    if signed {
        let half = 2f64.powi(8 * width as i32 - 1);
        let modulus = 2f64.powi(8 * width as i32);
        get(chunk, acc, line);
        num(chunk, half, line);
        ops::emit_dyn_lt(chunk, line);
        ops::emit_dyn_not(chunk, line);
        ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        get(chunk, acc, line);
        num(chunk, modulus, line);
        chunk.emit_op(Op::F64_SUB, line);
        set(chunk, acc, line);
        chunk.emit_end(line);
    }
    get(chunk, acc, line);
}

/// `Write(Boolean)` — one byte, 1 or 0.
pub fn emit_write_bool(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(2);
    let value = scratch;
    let bio = scratch + 1;
    set(chunk, value, line);
    set(chunk, bio, line);
    get(chunk, bio, line);
    get(chunk, value, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    num(chunk, 1.0, line);
    chunk.emit_else(line);
    num(chunk, 0.0, line);
    chunk.emit_end(line);
    emit_write_int(chunks, current, 1, line);
}

pub fn emit_read_bool(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_read_int(chunks, current, 1, false, line);
    let chunk = &mut chunks[current];
    num(chunk, 0.0, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
}

// ── floats ──────────────────────────────────────────────────────────────────

/// `Write(Single)` / `Write(Double)` — the IEEE-754 STORAGE bits, little-endian.
/// `bits.reinterpret_i32`/`_i64` is the same reinterpretation Fortran's
/// `TRANSFER` and Java's `floatToIntBits` reach.
/// ⛔ A DOUBLE'S 64 BITS DO NOT FIT IN A DOUBLE. The obvious spelling —
/// reinterpret to i64, convert to f64, hand that to `emit_write_int(8)` — is
/// silently lossy: `emit_write_int` peels bytes with f64 division, and an f64
/// carries 53 mantissa bits, so the low bits of the pattern are rounded away.
/// `2.718281828459` came back as `2.718281828458...` off byte 0. The 8-byte
/// form therefore splits the pattern into two 32-bit halves and never lets a
/// full 64-bit value touch an f64.
pub fn emit_write_float(chunks: &mut [Chunk], current: usize, width: u32, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(3);
    let value = scratch;
    let bio = scratch + 1;
    let bits = scratch + 2;
    set(chunk, value, line);
    set(chunk, bio, line);

    if width == 4 {
        get(chunk, bio, line);
        get(chunk, value, line);
        chunk.emit_op(Op::F32_DEMOTE_F64, line);
        chunk.emit_op(Op::I32_REINTERPRET_F32, line);
        chunk.emit_op(Op::F64_CONVERT_I32_U, line);
        emit_write_int(chunks, current, 4, line);
        return;
    }

    get(chunk, value, line);
    chunk.emit_op(Op::I64_REINTERPRET_F64, line);
    set(chunk, bits, line);

    // Little-endian: the low half goes first.
    get(chunk, bio, line);
    get(chunk, bits, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    emit_write_int(chunks, current, 4, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    get(chunk, bio, line);
    get(chunk, bits, line);
    chunk.emit_i64_const(32, line);
    chunk.emit_op(Op::I64_SHR_U, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    emit_write_int(chunks, current, 4, line);
}

pub fn emit_read_float(chunks: &mut [Chunk], current: usize, width: u32, line: u32) {
    if width == 4 {
        emit_read_int(chunks, current, 4, false, line);
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::I32_TRUNC_F64_U, line);
        chunk.emit_op(Op::F32_REINTERPRET_I32, line);
        chunk.emit_op(Op::F64_PROMOTE_F32, line);
        return;
    }

    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(3);
    let bio = scratch;
    let low = scratch + 1;
    let high = scratch + 2;
    set(chunk, bio, line);

    get(chunk, bio, line);
    emit_read_int(chunks, current, 4, false, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::I64_TRUNC_F64_U, line);
    set(chunk, low, line);

    get(chunk, bio, line);
    emit_read_int(chunks, current, 4, false, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::I64_TRUNC_F64_U, line);
    set(chunk, high, line);

    get(chunk, high, line);
    chunk.emit_i64_const(32, line);
    chunk.emit_op(Op::I64_SHL, line);
    get(chunk, low, line);
    chunk.emit_op(Op::I64_OR, line);
    chunk.emit_op(Op::F64_REINTERPRET_I64, line);
}

// ── strings and bytes ───────────────────────────────────────────────────────

/// The 7-bit encoded length prefix .NET puts in front of a string: seven bits
/// per byte, high bit set while more bytes follow.
fn emit_write_7bit(chunk: &mut Chunk, stream: u16, value: u16, line: u32) {
    let byte = chunk.alloc_scratch(1);
    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);

    get(chunk, value, line);
    num(chunk, 128.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    get(chunk, value, line);
    set(chunk, byte, line);
    put_byte(chunk, stream, byte, line);
    chunk.emit_br(2, line);
    chunk.emit_end(line);

    get(chunk, value, line);
    num(chunk, 128.0, line);
    fmod(chunk, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    num(chunk, 128.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, byte, line);
    put_byte(chunk, stream, byte, line);

    get(chunk, value, line);
    num(chunk, 128.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    set(chunk, value, line);
    chunk.emit_br(0, line);

    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
}

fn emit_read_7bit(chunk: &mut Chunk, stream: u16, out: u16, line: u32) {
    let scratch = chunk.alloc_scratch(3);
    let byte = scratch;
    let shift = scratch + 1;
    let acc = scratch + 2;
    num(chunk, 0.0, line);
    set(chunk, acc, line);
    num(chunk, 1.0, line);
    set(chunk, shift, line);

    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    take_byte(chunk, stream, byte, line);
    get(chunk, acc, line);
    get(chunk, byte, line);
    num(chunk, 128.0, line);
    fmod(chunk, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    get(chunk, shift, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, acc, line);
    get(chunk, byte, line);
    num(chunk, 128.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, shift, line);
    num(chunk, 128.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    set(chunk, shift, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
    get(chunk, acc, line);
    set(chunk, out, line);
}

pub fn emit_write_7bit_encoded_int(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(2);
    let value = scratch;
    let bio = scratch + 1;
    set(chunk, value, line);
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);
    emit_write_7bit(chunk, stream, value, line);
    core_wasm::undefined(chunk, line);
}

pub fn emit_read_7bit_encoded_int(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(2);
    let bio = scratch;
    let out = scratch + 1;
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);
    emit_read_7bit(chunk, stream, out, line);
    get(chunk, out, line);
}

/// `Write(String)` — 7-bit length prefix, then the UTF-8 bytes.
pub fn emit_write_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(5);
    let text = scratch;
    let bio = scratch + 1;
    let bytes = scratch + 2;
    let idx = scratch + 3;
    let byte = scratch + 4;
    set(chunk, text, line);
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);

    get(chunk, text, line);
    // WHATWG Encoding: `encode(encoder, text)` → UTF-8 bytes. Receiver-first,
    // and the encoder carries no state that matters here (UTF-8 always), so a
    // fresh one per call is correct rather than merely convenient.
    {
        let text_tmp = chunk.alloc_scratch(1);
        set(chunk, text_tmp, line);
        call(chunk, "web:encoding", "encoderNew", 0, line);
        get(chunk, text_tmp, line);
        call(chunk, "web:encoding", "encode", 2, line);
    }
    set(chunk, bytes, line);

    let len = chunk.alloc_scratch(1);
    get(chunk, bytes, line);
    call(chunk, "ecma:array", "length", 1, line);
    set(chunk, len, line);
    emit_write_7bit(chunk, stream, len, line);

    num(chunk, 0.0, line);
    set(chunk, idx, line);
    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    get(chunk, idx, line);
    get(chunk, bytes, line);
    call(chunk, "ecma:array", "length", 1, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, bytes, line);
    get(chunk, idx, line);
    // ⛔ Not `ecma:array.get`: `encode` hands back a Uint8Array, and that host
    // fn answers Undefined for anything that is not an Array or a Map.
    // `ARRAY_GET` is the VM's polymorphic index read.
    chunk.emit_op(Op::ARRAY_GET, line);
    set(chunk, byte, line);
    put_byte(chunk, stream, byte, line);
    get(chunk, idx, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, idx, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
    core_wasm::undefined(chunk, line);
}

pub fn emit_read_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(5);
    let bio = scratch;
    let count = scratch + 1;
    let bytes = scratch + 2;
    let idx = scratch + 3;
    let byte = scratch + 4;
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);
    emit_read_7bit(chunk, stream, count, line);

    call(chunk, "ecma:array", "new", 0, line);
    set(chunk, bytes, line);
    num(chunk, 0.0, line);
    set(chunk, idx, line);
    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    get(chunk, idx, line);
    get(chunk, count, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    take_byte(chunk, stream, byte, line);
    get(chunk, bytes, line);
    get(chunk, byte, line);
    call(chunk, "ecma:array", "push", 2, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, idx, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, idx, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);

    get(chunk, bytes, line);
    // ⛔ WHATWG `TextDecoder.decode` takes a BufferSource. A plain Array
    // decodes to the EMPTY STRING rather than failing, so the conversion is
    // not optional.
    call(chunk, "ecma:uint8array", "newFromIterable", 1, line);
    {
        let bytes_tmp = chunk.alloc_scratch(1);
        set(chunk, bytes_tmp, line);
        call(chunk, "web:encoding", "decoderNew", 0, line);
        get(chunk, bytes_tmp, line);
        call(chunk, "web:encoding", "decode", 2, line);
    }
}

/// `Write(Byte())` — the raw bytes, no length prefix. That asymmetry with
/// `Write(String)` is .NET's, and it is why `ReadBytes(n)` takes a count.
pub fn emit_write_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(5);
    let src = scratch;
    let bio = scratch + 1;
    let idx = scratch + 2;
    let byte = scratch + 3;
    set(chunk, src, line);
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);

    num(chunk, 0.0, line);
    set(chunk, idx, line);
    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    get(chunk, idx, line);
    get(chunk, src, line);
    call(chunk, "ecma:array", "length", 1, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, src, line);
    get(chunk, idx, line);
    // ⛔ Not `ecma:array.get`: the source is whatever the caller passed —
    // a VB `Byte()`, but also the Uint8Array `TextEncoder.encode` hands back
    // for `Write(Char)`/`Write(String)` — and that host fn answers Undefined
    // for anything that is not an Array or a Map.
    chunk.emit_op(Op::ARRAY_GET, line);
    set(chunk, byte, line);
    put_byte(chunk, stream, byte, line);
    get(chunk, idx, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, idx, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
    core_wasm::undefined(chunk, line);
}

/// `ReadBytes(count)` — up to `count` bytes; a short read is NOT an error here,
/// which is what separates it from the fixed-width readers.
pub fn emit_read_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(5);
    let count = scratch;
    let bio = scratch + 1;
    let out = scratch + 2;
    let idx = scratch + 3;
    set(chunk, count, line);
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);

    call(chunk, "ecma:array", "new", 0, line);
    set(chunk, out, line);
    num(chunk, 0.0, line);
    set(chunk, idx, line);
    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    get(chunk, idx, line);
    get(chunk, count, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    field(chunk, stream, POS, line);
    field(chunk, stream, LEN, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    field(chunk, stream, BUF, line);
    field(chunk, stream, POS, line);
    call(chunk, "ecma:array", "get", 2, line);
    let byte = chunk.alloc_scratch(1);
    set(chunk, byte, line);
    field(chunk, stream, POS, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    store(chunk, stream, POS, line);
    get(chunk, out, line);
    get(chunk, byte, line);
    call(chunk, "ecma:array", "push", 2, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, idx, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, idx, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
    get(chunk, out, line);
}

/// `Read(buffer, index, count)` — fills the CALLER'S array and answers how
/// many bytes it actually got.
///
/// ⛔ Distinct from `Read(count)`/`ReadBytes(count)`, which allocate and return
/// a new array. .NET declares both, separated only by arity, and only the
/// one-argument form was registered — so `br.Read(buffer, 0, 5)` resolved to
/// nothing, left `buffer` untouched and answered no value.
///
/// A short read is not an error: at end of stream the loop simply stops, and
/// the returned count is what the caller checks.
pub fn emit_read_into_buffer(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(6);
    let count = scratch;
    let index = scratch + 1;
    let buffer = scratch + 2;
    let bio = scratch + 3;
    let read = scratch + 4;
    let byte = scratch + 5;
    set(chunk, count, line);
    set(chunk, index, line);
    set(chunk, buffer, line);
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);

    num(chunk, 0.0, line);
    set(chunk, read, line);
    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);

    get(chunk, read, line);
    get(chunk, count, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);

    field(chunk, stream, POS, line);
    field(chunk, stream, LEN, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);

    field(chunk, stream, BUF, line);
    field(chunk, stream, POS, line);
    call(chunk, "ecma:array", "get", 2, line);
    set(chunk, byte, line);
    field(chunk, stream, POS, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    store(chunk, stream, POS, line);

    get(chunk, buffer, line);
    get(chunk, index, line);
    get(chunk, read, line);
    chunk.emit_op(Op::F64_ADD, line);
    get(chunk, byte, line);
    call(chunk, "ecma:array", "set", 3, line);
    chunk.emit_op(Op::DROP, line);

    get(chunk, read, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, read, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);

    get(chunk, read, line);
}

/// `Write(Char)` — the UTF-8 bytes, with NO length prefix.
pub fn emit_write_char(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(2);
    let ch = scratch;
    let bio = scratch + 1;
    set(chunk, ch, line);
    set(chunk, bio, line);
    get(chunk, bio, line);
    get(chunk, ch, line);
    // WHATWG Encoding: `encode(encoder, text)` → UTF-8 bytes. Receiver-first,
    // and the encoder carries no state that matters here (UTF-8 always), so a
    // fresh one per call is correct rather than merely convenient.
    {
        let text_tmp = chunk.alloc_scratch(1);
        set(chunk, text_tmp, line);
        call(chunk, "web:encoding", "encoderNew", 0, line);
        get(chunk, text_tmp, line);
        call(chunk, "web:encoding", "encode", 2, line);
    }
    emit_write_bytes(chunks, current, line);
}

pub fn emit_read_char(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_read_int(chunks, current, 1, false, line);
    let chunk = &mut chunks[current];
    shared_strings::emit_from_char_code(chunk, line);
}

/// `PeekChar` — the next character WITHOUT advancing, or -1 at the end.
pub fn emit_peek_char(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let bio = chunk.alloc_scratch(1);
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);
    field(chunk, stream, POS, line);
    field(chunk, stream, LEN, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    field(chunk, stream, BUF, line);
    field(chunk, stream, POS, line);
    call(chunk, "ecma:array", "get", 2, line);
    chunk.emit_else(line);
    num(chunk, -1.0, line);
    chunk.emit_end(line);
}

// ── the rest of the surface ─────────────────────────────────────────────────

/// `Seek(offset, origin)` on the writer moves the underlying stream's cursor —
/// which is what makes a later `Write` OVERWRITE rather than append.
pub fn emit_seek(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(3);
    let origin = scratch;
    let offset = scratch + 1;
    let bio = scratch + 2;
    set(chunk, origin, line);
    set(chunk, offset, line);
    set(chunk, bio, line);
    let stream = stream_of(chunk, bio, line);

    // 0 Begin, 1 Current, 2 End.
    get(chunk, origin, line);
    num(chunk, 1.0, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    field(chunk, stream, POS, line);
    get(chunk, offset, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    get(chunk, origin, line);
    num(chunk, 2.0, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    field(chunk, stream, LEN, line);
    get(chunk, offset, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    get(chunk, offset, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    store(chunk, stream, POS, line);
    field(chunk, stream, POS, line);
}

/// `Flush` — nothing is buffered above the stream, so this is observable only
/// in that it must not fail.
pub fn emit_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    core_wasm::undefined(chunk, line);
}

/// `Close` / `Dispose` — closes the underlying stream unless the constructor
/// was told to leave it open, which is the whole point of that third argument.
pub fn emit_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let bio = chunk.alloc_scratch(1);
    set(chunk, bio, line);
    field(chunk, bio, LEAVE_OPEN, line);
    ops::emit_dyn_to_bool(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    let stream = stream_of(chunk, bio, line);
    get(chunk, stream, line);
    core_wasm::bool_const(chunk, line, true);
    field_set(chunk, "__ms_closed", line);
    chunk.emit_end(line);
    core_wasm::undefined(chunk, line);
}

/// `Write(value)` with no static width — .NET's default for each shape.
///
/// This is the half of the overload that RUNTIME can answer: a string, a
/// boolean and an array are distinguishable at run time, and a number with a
/// fractional part is a `Double` where a whole one is an `Int32` — which is
/// exactly what .NET picks for an unsuffixed literal. What runtime cannot
/// answer is `Int16` vs `Int32` vs `UInt64`: those are all whole numbers here,
/// and the VB walker routes them to the width-specific spellings instead.
pub fn emit_write_auto(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(3);
    let value = scratch;
    let bio = scratch + 1;
    let kind = scratch + 2;
    set(chunk, value, line);
    set(chunk, bio, line);

    get(chunk, value, line);
    call(chunk, "ecma:value", "typeof", 1, line);
    set(chunk, kind, line);

    // string
    get(chunk, kind, line);
    chunk.emit_string_const("string", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    get(chunk, bio, line);
    get(chunk, value, line);
    emit_write_string(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);

    // boolean
    get(chunk, kind, line);
    chunk.emit_string_const("boolean", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    get(chunk, bio, line);
    get(chunk, value, line);
    emit_write_bool(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);

    // Byte() — the array overload writes the bytes raw.
    get(chunk, value, line);
    call(chunk, "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    get(chunk, bio, line);
    get(chunk, value, line);
    emit_write_bytes(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);

    // A whole number is an Int32, anything else a Double.
    get(chunk, value, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    get(chunk, bio, line);
    get(chunk, value, line);
    emit_write_int(chunks, current, 4, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);
    get(chunk, bio, line);
    get(chunk, value, line);
    emit_write_float(chunks, current, 8, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    core_wasm::undefined(chunk, line);
}
