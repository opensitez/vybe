use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;
use vybe_compiler::primitives::instructions::{core_wasm, host};

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
    chunk.emit_op_u16(Op::CALL_IMPORT, to_f64, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, number_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    chunk.emit_f64_const(2_147_483_647.0, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    for _ in 0..7 {
        chunk.emit_i32_const(0, line);
    }
    chunk.emit_array_new_fixed(0, 8, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    for _ in 0..3 {
        chunk.emit_i32_const(0, line);
    }
    chunk.emit_array_new_fixed(0, 4, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_to_number(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let (bytes_slot, offset_slot) = stash_two(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, offset_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

pub fn emit_to_boolean(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_to_number(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_to_char(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_to_number(chunks, current, line);
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
    chunk.emit_op(Op::DROP, line);

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
