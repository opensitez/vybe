use vybe_compiler::primitives::instructions::host;
use vybe_runtime::{opcode::Op, Chunk};
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

    // This used to emit `[value, 0, 0, 0]` -- the whole number in element 0
    // with zero padding, not its BYTES. It round-tripped only because the
    // reader was the matching stub (one `ARRAY_GET` at element 0), so the two
    // wrongs cancelled and `GetBytes(300)(0)` answered 300 instead of 44.
    // Both sides are on `primitives::packing` now.
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    // Magnitude, not value: `-9e18` is not `> i32::MAX`, so a bare `>` test
    // sent it down the 4-byte lane and threw the top bytes away.
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_op(Op::F64_ABS, line);
    chunk.emit_f64_const(2_147_483_647.0, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    // The 64-bit / Double lane is still a paired stub: the value itself in
    // element 0, zero-padded, read back by `emit_to_number64`. A real encoding
    // needs the IEEE-754 bit pattern (`bits::emit_reinterpret` to `I64`) for a
    // Double and true 8-byte integer bytes for a Long, and the two cannot be
    // told apart here: `GetBytes` picks its overload from the argument's
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

fn emit_bytes_of_width(
    chunks: &mut [Chunk],
    current: usize,
    number_slot: u16,
    width: i32,
    line: u32,
) {
    use vybe_compiler::primitives::packing::{
        emit_store_u32_to_array_from_number_slot, emit_store_u64_parts_to_array_from_number_slots,
        Endian,
    };
    for _ in 0..width {
        chunks[current].emit_i32_const(0, line);
    }
    chunks[current].emit_array_new_fixed(0, width as u16, line);
    let array_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    // The packing stores narrow with `I32_TRUNC_F64_U`, which traps on a
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
