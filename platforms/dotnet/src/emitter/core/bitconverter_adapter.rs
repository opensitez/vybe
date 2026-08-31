use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn stash_two(chunk: &mut Chunk, line: u32) -> (u16, u16) {
    let first = chunk.alloc_scratch(2);
    let second = first + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, second, line);
    chunk.emit_op_u16(Op::LOCAL_SET, first, line);
    (first, second)
}

pub fn emit_get_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.alloc_scratch(2);
    let number_slot = value_slot + 1;
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");

    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "wasm:js-boolean", "test", 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    chunk.emit_array_new_fixed(0, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "wasm:js-string", "test", 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_i32_const(0, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_array_new_fixed(0, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, number_slot, line);

    // ⛔ THIS USED TO EMIT `[value, 0, 0, 0]` — the whole number in element 0
    // with zero padding, not its BYTES. It round-tripped only because the
    // reader was the matching stub (one `ARRAY_GET` at element 0), so the two
    // wrongs cancelled and `GetBytes(300)(0)` answered 300 instead of 44.
    // Both sides are on `primitives::packing` now.
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    // ⛔ MAGNITUDE, not value: `-9e18` is not `> i32::MAX`, so a bare `>` test
    // sent it down the 4-byte lane and threw the top bytes away.
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_op(Op::F64_ABS, line);
    chunk.emit_f64_const(2_147_483_647.0, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    // ⚠ THE 64-BIT / DOUBLE LANE IS STILL A PAIRED STUB: the value itself in
    // element 0, zero-padded, read back by `emit_to_number64`. A real encoding
    // needs the IEEE-754 bit pattern (`bits::emit_reinterpret` to `I64`) for a
    // Double and true 8-byte integer bytes for a Long, and the two cannot be
    // told apart here — `GetBytes` picks its overload from the argument's
    // STATIC type in .NET, and this adapter sees only a runtime f64. Converting
    // one side alone regresses the round-trips (measured: 7 tests, including
    // csharp `bit_converter_double_bytes_reconstruct_original_fraction`).
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    for _ in 0..7 {
        chunks[current].emit_i32_const(0, line);
    }
    chunks[current].emit_array_new_fixed(0, 8, line);
    chunks[current].emit_else(line);
    emit_bytes_of_width(chunks, current, number_slot, 4, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// The 64-bit/double reader — the OTHER half of the stub above. Reads element
/// `startIndex` whole, because that is where its writer put the value.
pub fn emit_to_number64(chunks: &mut [Chunk], current: usize, line: u32) {
    let (bytes_slot, offset_slot) = stash_two(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, offset_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
}

/// A zero-filled byte array of `width`, with `number_slot`'s little-endian
/// bytes written into it by [`packing`]. Stack: `[] -> [array]`.
fn emit_bytes_of_width(
    chunks: &mut [Chunk],
    current: usize,
    number_slot: u16,
    width: i32,
    line: u32,
) {
    use vybe_compiler::primitives::packing::{
        Endian, emit_store_u32_to_array_from_number_slot,
        emit_store_u64_parts_to_array_from_number_slots,
    };
    for _ in 0..width {
        chunks[current].emit_i32_const(0, line);
    }
    chunks[current].emit_array_new_fixed(0, width as u16, line);
    let array_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    // ⛔ The packing stores narrow with `I32_TRUNC_F64_U`, which TRAPS on a
    // negative. Two's complement IS the unsigned value `2^32 + n`, so
    // normalising here is both what the encoding means and what keeps the
    // shared helper total.
    let number_slot = {
        let unsigned = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, number_slot, line);
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_op(Op::F64_LT, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, number_slot, line);
        chunks[current].emit_f64_const(4_294_967_296.0, line);
        chunks[current].emit_op(Op::F64_ADD, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, number_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, unsigned, line);
        unsigned
    };
    if width == 8 {
        // The high word is `value / 2^32`, the low word `value mod 2^32`.
        let hi_slot = chunks[current].alloc_scratch(1);
        let lo_slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, number_slot, line);
        chunks[current].emit_f64_const(4_294_967_296.0, line);
        chunks[current].emit_op(Op::F64_DIV, line);
        chunks[current].emit_op(Op::F64_FLOOR, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, hi_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, number_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, hi_slot, line);
        chunks[current].emit_f64_const(4_294_967_296.0, line);
        chunks[current].emit_op(Op::F64_MUL, line);
        chunks[current].emit_op(Op::F64_SUB, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, lo_slot, line);
        emit_store_u64_parts_to_array_from_number_slots(
            chunks,
            current,
            array_slot,
            hi_slot,
            lo_slot,
            Endian::Little,
            line,
        );
    } else {
        emit_store_u32_to_array_from_number_slot(
            chunks,
            current,
            array_slot,
            number_slot,
            Endian::Little,
            line,
        );
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
}

/// `BitConverter.To*(bytes, startIndex)` — the shared little-endian byte
/// decode from [`packing`], which already owns this for Lua `string.unpack`,
/// PHP `unpack` and Ruby `String#unpack`.
///
/// ⛔ `width` BYTES, not one. Reading a single `ARRAY_GET` at `startIndex` is
/// the low byte only: it answers correctly for any value under 256 and silently
/// truncates everything else, which is why `ToInt32` looked fine on small
/// numbers. .NET fixes the width per method name, so each leaf states its own.
///
/// `startIndex` is a runtime value while the packing helpers index from 0, so
/// the window is sliced first — one `ecma:array.slice`, on arrays of 2 to 8
/// bytes.
fn emit_to_width(chunks: &mut [Chunk], current: usize, width: i32, line: u32) {
    let (bytes_slot, offset_slot) = stash_two(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, offset_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, offset_slot, line);
    chunks[current].emit_i32_const(width, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    vybe_compiler::primitives::collections::emit_slice(chunks, current, line);
    let window_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, window_slot, line);
    use vybe_compiler::primitives::packing::{
        Endian, emit_load_u16_from_array_f64, emit_load_u32_from_array_f64,
    };
    if width <= 2 {
        emit_load_u16_from_array_f64(chunks, current, window_slot, Endian::Little, line);
    } else {
        emit_load_u32_from_array_f64(chunks, current, window_slot, Endian::Little, line);
    }
}

/// Reinterpret an unsigned little-endian read as two's-complement signed:
/// anything at or above half the range is that value minus the full range.
fn emit_sign_extend(chunks: &mut [Chunk], current: usize, bits: i32, line: u32) {
    let half = 2f64.powi(bits - 1);
    let full = 2f64.powi(bits);
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_f64_const(half, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_f64_const(full, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_end(line);
}

/// `ToUInt32` — a four-byte window, unsigned.
pub fn emit_to_number(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_to_width(chunks, current, 4, line);
}

/// `ToInt32` — the same window, two's-complement.
pub fn emit_to_int32(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_to_width(chunks, current, 4, line);
    emit_sign_extend(chunks, current, 32, line);
}

/// `ToUInt16`/`ToChar` — a two-byte window, unsigned.
pub fn emit_to_number16(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_to_width(chunks, current, 2, line);
}

/// `ToInt16` — the same window, two's-complement.
pub fn emit_to_int16(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_to_width(chunks, current, 2, line);
    emit_sign_extend(chunks, current, 16, line);
}

/// `ToBoolean` reads ONE byte — .NET defines it as "nonzero at startIndex".
pub fn emit_to_boolean(chunks: &mut [Chunk], current: usize, line: u32) {
    let (bytes_slot, offset_slot) = stash_two(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, offset_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// A .NET `Char` is UTF-16: TWO bytes, not one.
pub fn emit_to_char(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_to_number16(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:string", "fromCharCode", 1, line);
}

pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let bytes_slot = chunk.alloc_scratch(5);
    let len_slot = bytes_slot + 1;
    let i_slot = bytes_slot + 2;
    let result_slot = bytes_slot + 3;
    let part_slot = bytes_slot + 4;

    chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let block = chunk.emit_block(line);
    let (loop_pos, _) = chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_i32_const(16, line);
    host::emit(chunk, "ecma:number", "toString", 2, line);
    chunk.emit_i32_const(2, line);
    chunk.emit_string_const("0", line);
    host::emit(chunk, "ecma:string", "padStart", 3, line);
    host::emit(chunk, "ecma:string", "toUpperCase", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, part_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_string_const("-", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, part_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_pos);
    chunk.emit_end(line);
    chunk.patch_block(block);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_block_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let src_slot = chunk.alloc_scratch(6);
    let src_offset_slot = src_slot + 1;
    let dst_slot = src_slot + 2;
    let dst_offset_slot = src_slot + 3;
    let count_slot = src_slot + 4;
    let i_slot = src_slot + 5;

    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dst_offset_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dst_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, src_offset_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, src_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let block = chunk.emit_block(line);
    let (loop_pos, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, dst_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, dst_offset_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, src_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, src_offset_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_pos);
    chunk.emit_end(line);
    chunk.patch_block(block);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_is_little_endian(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::bool_const(&mut chunks[current], line, true);
}

// ── Bit reinterpretation ────────────────────────────────────────────────────
//
// `System.BitConverter`'s bit-cast family. These lower to the SAME shared
// emitter the AST's `UnaryOp::Reinterpret` uses — the node Fortran's `TRANSFER`,
// Go's `Float32bits` and Java's `floatToIntBits` all reach — so the concept
// stays unified while the .NET SPELLING lives here, in the platform that owns
// `System.*`, rather than in a language walker.
//
// Stack: `[value] -> [reinterpreted]`.

pub fn emit_single_to_int32_bits(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::bits::emit_reinterpret(
        &mut chunks[current],
        vybe_ast::NumericRepr::I32,
        line,
    );
}

pub fn emit_int32_bits_to_single(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::bits::emit_reinterpret(
        &mut chunks[current],
        vybe_ast::NumericRepr::F32,
        line,
    );
}

pub fn emit_double_to_int64_bits(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::bits::emit_reinterpret(
        &mut chunks[current],
        vybe_ast::NumericRepr::I64,
        line,
    );
}

pub fn emit_int64_bits_to_double(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::bits::emit_reinterpret(
        &mut chunks[current],
        vybe_ast::NumericRepr::F64,
        line,
    );
}
