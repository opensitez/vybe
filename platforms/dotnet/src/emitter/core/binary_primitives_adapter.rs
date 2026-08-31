//! `System.Buffers.Binary.BinaryPrimitives` — byte-order reads and writes.
//!
//! ## Arithmetic, not shifts — deliberately
//!
//! The obvious lowering is `b0 | b1 << 8 | …`. It is wrong here twice over:
//! shifts on this runtime are i32 and WRAP mod 32 (measured: `1UL << 50` is
//! `262144`), so the 64-bit forms would silently fold, and the i32 sign bit
//! would corrupt the top byte of a `UInt32`. Multiplying by `256^i` in f64 is
//! exact for every value an f64 can hold and has neither failure mode.
//!
//! ⛔ CEILING, STATED: values are f64, so the 64-bit forms are exact only below
//! 2^53. That is the same representation gap that makes `1UL << 50` wrong, and
//! it is not something this adapter can fix — see the numeric-literal notes.
//!
//! `DataView` was the other candidate and does not apply: `new byte[4]` is a
//! plain indexed array here, not a `Uint8Array` over an `ArrayBuffer`, so there
//! is no buffer to view. Measured before choosing.

use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn push_f64(chunk: &mut Chunk, value: f64, line: u32) {
    chunk.emit_f64_const(value, line);
}

/// `256^n`, as the exact f64 it always is for n ≤ 8.
fn place(n: u32) -> f64 {
    256f64.powi(n as i32)
}

/// One element read: `array[index]`.
fn index_get(chunk: &mut Chunk, array: u16, index: u32, line: u32) {
    get(chunk, array, line);
    push_f64(chunk, index as f64, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// `floor(value_slot / divisor) - floor(value_slot / (divisor * 256)) * 256`
/// — the byte at position `divisor`, without any shift.
fn push_byte_at(chunk: &mut Chunk, value: u16, divisor: f64, line: u32) {
    get(chunk, value, line);
    push_f64(chunk, divisor, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    get(chunk, value, line);
    push_f64(chunk, divisor * 256.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    push_f64(chunk, 256.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
}

/// `Read{I,U}{16,32,64}{Little,Big}Endian(source)` — stack `[src]` → `[value]`.
pub fn emit_read(
    chunks: &mut [Chunk],
    current: usize,
    width: u32,
    signed: bool,
    little: bool,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let base = chunk.alloc_scratch(2);
    let (src, acc) = (base, base + 1);
    set(chunk, src, line);
    push_f64(chunk, 0.0, line);
    set(chunk, acc, line);

    for i in 0..width {
        // Byte `i` of the VALUE sits at index `i` little-endian, and at the
        // mirrored index big-endian. Only the index differs.
        let index = if little { i } else { width - 1 - i };
        get(chunk, acc, line);
        index_get(chunk, src, index, line);
        push_f64(chunk, place(i), line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        set(chunk, acc, line);
    }

    if signed {
        // Two's complement: anything at or above half the range is negative.
        let span = place(width);
        let half = span / 2.0;
        get(chunk, acc, line);
        push_f64(chunk, half, line);
        ops::emit_dyn_ge(chunk, line);
        chunk.emit_if(line);
        get(chunk, acc, line);
        push_f64(chunk, span, line);
        chunk.emit_op(Op::F64_SUB, line);
        set(chunk, acc, line);
        chunk.emit_end(line);
    }
    get(chunk, acc, line);
}

/// `Write{I,U}{16,32,64}{Little,Big}Endian(destination, value)` —
/// stack `[dest, value]` → `[null]`.
pub fn emit_write(
    chunks: &mut [Chunk],
    current: usize,
    width: u32,
    little: bool,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let base = chunk.alloc_scratch(2);
    let (value, dest) = (base, base + 1);
    set(chunk, value, line);
    set(chunk, dest, line);

    // ⛔ NORMALISE FIRST. A negative `int` has to become its unsigned twin
    // before any byte can be extracted — `floor(-1 / 256)` is `-1`, not `255`.
    let span = place(width);
    get(chunk, value, line);
    push_f64(chunk, 0.0, line);
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    get(chunk, value, line);
    push_f64(chunk, span, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, value, line);
    chunk.emit_end(line);

    for i in 0..width {
        let index = if little { i } else { width - 1 - i };
        get(chunk, dest, line);
        push_f64(chunk, index as f64, line);
        push_byte_at(chunk, value, place(i), line);
        chunk.emit_op(Op::ARRAY_SET, line);
    }
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `ReverseEndianness(value)` — read the bytes out and back in mirrored.
pub fn emit_reverse_endianness(
    chunks: &mut [Chunk],
    current: usize,
    width: u32,
    signed: bool,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let base = chunk.alloc_scratch(2);
    let (value, acc) = (base, base + 1);
    set(chunk, value, line);

    let span = place(width);
    get(chunk, value, line);
    push_f64(chunk, 0.0, line);
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    get(chunk, value, line);
    push_f64(chunk, span, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, value, line);
    chunk.emit_end(line);

    push_f64(chunk, 0.0, line);
    set(chunk, acc, line);
    for i in 0..width {
        get(chunk, acc, line);
        push_byte_at(chunk, value, place(i), line);
        push_f64(chunk, place(width - 1 - i), line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        set(chunk, acc, line);
    }
    if signed {
        let half = span / 2.0;
        get(chunk, acc, line);
        push_f64(chunk, half, line);
        ops::emit_dyn_ge(chunk, line);
        chunk.emit_if(line);
        get(chunk, acc, line);
        push_f64(chunk, span, line);
        chunk.emit_op(Op::F64_SUB, line);
        set(chunk, acc, line);
        chunk.emit_end(line);
    }
    get(chunk, acc, line);
}

/// Decode `read_i32_le` / `write_u16_be` / `reverse_i64` into its parameters.
pub fn parse_key(suffix: &str) -> Option<(&'static str, u32, bool, bool)> {
    let mut parts = suffix.split('_');
    let op = match parts.next()? {
        "read" => "read",
        "write" => "write",
        "reverse" => "reverse",
        _ => return None,
    };
    let ty = parts.next()?;
    let signed = ty.starts_with('i');
    let width = match &ty[1..] {
        "16" => 2,
        "32" => 4,
        "64" => 8,
        _ => return None,
    };
    let little = match parts.next() {
        Some("le") => true,
        Some("be") => false,
        None if op == "reverse" => true,
        _ => return None,
    };
    Some((op, width, signed, little))
}

/// `MemoryMarshal.Cast<TFrom, TTo>(source)` — reinterpret the bytes.
///
/// ⛔ THE TYPE ARGUMENTS ARE THE WHOLE OPERATION and the emitter cannot see
/// them, so the C# walker rewrites the generic call into `CastBytes(src, from,
/// to)` carrying the two widths — the same route `Unsafe.SizeOf<T>` takes, and
/// for the same reason: a framework generic's arguments are not passed at
/// runtime, so anything that depends on them folds where they still exist.
///
/// Bytes are little-endian, matching every other reinterpret here, and the
/// element count truncates the way .NET's does: a 3-byte source cast to `int`
/// yields ZERO elements, not a partial one.
pub fn emit_cast(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 3..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let base = chunk.alloc_scratch(9);
    let (to_w, from_w, src, bytes, out, i, j, acc, n) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
        base + 7,
        base + 8,
    );
    set(chunk, to_w, line);
    set(chunk, from_w, line);
    set(chunk, src, line);

    // Flatten to bytes: each source element contributes `from_w` of them.
    let new_array = chunk.add_import("ecma:array", "new");
    chunk.emit_call(new_array, 0, line);
    set(chunk, bytes, line);
    let arr_len = chunk.add_import("ecma:array", "length");
    let push = chunk.add_import("ecma:array", "push");
    get(chunk, src, line);
    chunk.emit_call(arr_len, 1, line);
    set(chunk, n, line);
    push_f64(chunk, 0.0, line);
    set(chunk, i, line);
    let g1 = chunk.emit_block(line);
    let b1 = chunk.emit_block(line);
    let (l1, _) = chunk.emit_loop_s(line);
    get(chunk, i, line);
    get(chunk, n, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    push_f64(chunk, 0.0, line);
    set(chunk, j, line);
    let g2 = chunk.emit_block(line);
    let b2 = chunk.emit_block(line);
    let (l2, _) = chunk.emit_loop_s(line);
    get(chunk, j, line);
    get(chunk, from_w, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, bytes, line);
    // `floor(src[i] / 256^j) mod 256`
    get(chunk, src, line);
    get(chunk, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    set(chunk, acc, line);
    push_byte_at_dynamic(chunk, acc, j, line);
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, j, line);
    push_f64(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, j, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(l2);
    chunk.emit_end(line);
    chunk.patch_block(b2);
    chunk.emit_end(line);
    chunk.patch_block(g2);
    get(chunk, i, line);
    push_f64(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(l1);
    chunk.emit_end(line);
    chunk.patch_block(b1);
    chunk.emit_end(line);
    chunk.patch_block(g1);

    // Regroup into `to_w`-wide values, truncating any partial tail.
    chunk.emit_call(new_array, 0, line);
    set(chunk, out, line);
    get(chunk, bytes, line);
    chunk.emit_call(arr_len, 1, line);
    get(chunk, to_w, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    set(chunk, n, line);
    push_f64(chunk, 0.0, line);
    set(chunk, i, line);
    let g3 = chunk.emit_block(line);
    let b3 = chunk.emit_block(line);
    let (l3, _) = chunk.emit_loop_s(line);
    get(chunk, i, line);
    get(chunk, n, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    push_f64(chunk, 0.0, line);
    set(chunk, acc, line);
    push_f64(chunk, 0.0, line);
    set(chunk, j, line);
    let g4 = chunk.emit_block(line);
    let b4 = chunk.emit_block(line);
    let (l4, _) = chunk.emit_loop_s(line);
    get(chunk, j, line);
    get(chunk, to_w, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, acc, line);
    get(chunk, bytes, line);
    get(chunk, i, line);
    get(chunk, to_w, line);
    chunk.emit_op(Op::F64_MUL, line);
    get(chunk, j, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    push_pow256(chunk, j, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, acc, line);
    get(chunk, j, line);
    push_f64(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, j, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(l4);
    chunk.emit_end(line);
    chunk.patch_block(b4);
    chunk.emit_end(line);
    chunk.patch_block(g4);
    get(chunk, out, line);
    get(chunk, acc, line);
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, i, line);
    push_f64(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(l3);
    chunk.emit_end(line);
    chunk.patch_block(b3);
    chunk.emit_end(line);
    chunk.patch_block(g3);
    get(chunk, out, line);
}

/// `256^n` where `n` is a runtime slot — repeated doubling would need a loop,
/// so this uses `exp(n * ln 256)` via the host `pow`.
fn push_pow256(chunk: &mut Chunk, n: u16, line: u32) {
    push_f64(chunk, 256.0, line);
    get(chunk, n, line);
    let pow = chunk.add_import("ecma:math", "pow");
    chunk.emit_call(pow, 2, line);
}

/// `floor(value / 256^j) mod 256` with a runtime `j`.
fn push_byte_at_dynamic(chunk: &mut Chunk, value: u16, j: u16, line: u32) {
    get(chunk, value, line);
    push_pow256(chunk, j, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    let tmp = chunk.alloc_scratch(1);
    set(chunk, tmp, line);
    get(chunk, tmp, line);
    get(chunk, tmp, line);
    push_f64(chunk, 256.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    push_f64(chunk, 256.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
}
