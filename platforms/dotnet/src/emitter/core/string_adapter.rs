use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_emitter::instructions::{core_wasm, host};

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

pub fn emit_string_from_chars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_string_const("", line);
        return;
    }
    let chars_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, chars_slot, line);
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }

    let string_test_idx = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, string_test_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunks[current].emit_else(line);
    emit_char_array_to_string(chunks, current, chars_slot, line);
    chunks[current].emit_end(line);
}

fn emit_char_array_to_string(chunks: &mut [Chunk], current: usize, chars_slot: u16, line: u32) {
    let char_code_idx = chunks[current].add_import("wasm:js-string", "charCodeAt");
    let from_chars_idx = chunks[current].add_import("wasm:js-string", "fromCharCodeArray");
    let string_test_idx = chunks[current].add_import("wasm:js-string", "test");
    let chunk = &mut chunks[current];
    let units_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    let elem_slot = reserve_slot(chunk);

    vybe_emitter::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, units_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_emitter::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_emitter::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, string_test_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    host::emit(&mut chunks[current], "ecma:string", "String", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, char_code_idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_end(line);
    vybe_emitter::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_emitter::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, from_chars_idx, line);
    chunks[current].emit(3, line);
}
