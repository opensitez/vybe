use vybe_runtime::{Chunk, Op};
fn alloc_local(c: &mut Chunk) -> u16 {
    c.alloc_scratch(1)
}
fn lget(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn lset(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
fn emit_import_call(cs: &mut [Chunk], current: usize, m: &str, n: &str, argc: u8, l: u32) {
    let i = cs[current].add_import(m, n);
    cs[current].emit_call(i, argc, l);
}
pub fn emit_rank(chunks: &mut [Chunk], current: usize, line: u32) {
    fn emit_is_array(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        emit_import_call(chunks, current, "ecma:array", "isArray", 1, line);
        emit_import_call(chunks, current, "wasm:js-boolean", "cast", 1, line);
    }

    fn emit_array_get_zero(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_i32_const(0, line);
        emit_import_call(chunks, current, "ecma:array", "get", 2, line);
    }

    let value_slot = alloc_local(&mut chunks[current]);
    let first_slot = alloc_local(&mut chunks[current]);
    let second_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], value_slot, line);
    emit_is_array(chunks, current, value_slot, line);
    chunks[current].emit_if_value(line);

    emit_array_get_zero(chunks, current, value_slot, line);
    lset(&mut chunks[current], first_slot, line);
    emit_is_array(chunks, current, first_slot, line);
    chunks[current].emit_if_value(line);

    emit_array_get_zero(chunks, current, first_slot, line);
    lset(&mut chunks[current], second_slot, line);
    emit_is_array(chunks, current, second_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
}

pub fn emit_pop(chunks: &mut [Chunk], current: usize, line: u32) {
    let src = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, src, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    emit_import_call(chunks, current, "wasm:js-boolean", "cast", 1, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    emit_import_call(chunks, current, "ecma:array", "pop", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_else(line);

    // Keyed: take the last insertion-ordered key, read it, drop it.
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    emit_import_call(chunks, current, "ecma:map", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, keys, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    emit_import_call(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    emit_import_call(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}
