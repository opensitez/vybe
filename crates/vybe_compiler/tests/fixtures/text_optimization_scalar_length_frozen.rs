use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::{Chunk, Op};
fn alloc_local(c: &mut Chunk) -> u16 {
    c.alloc_scratch(1)
}
fn lget(c: &mut Chunk, s: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, line);
}
fn lset(c: &mut Chunk, s: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, line);
}
fn emit_import_call(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}
pub fn emit_len(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], value_slot, line);

    lget(&mut chunks[current], value_slot, line);
    {
        let idx = chunks[current].add_import("wasm:js-string", "test");
        chunks[current].emit_call(idx, 1, line);
    }
    // wasm:js-string.test already returns I32(0/1) — use it directly as the if
    // condition. Do NOT call emit_dyn_to_bool here: that registers imports on
    // chunks[current] via chunk.add_import, which collides with the global
    // import indices emitted by emit_import_call (chunks[0]-based) below,
    // causing CALL_IMPORT to resolve the wrong host function at runtime.
    chunks[current].emit_if_value(line);

    // String — wasm:js-string.length.
    lget(&mut chunks[current], value_slot, line);
    emit_import_call(chunks, current, "wasm:js-string", "length", 1, line);

    chunks[current].emit_else(line);

    // Not a string — try array length first. A null result means the
    // value wasn't array-like; a numeric 0 is a real empty length and
    // must not fall through to Object.keys(...).length.
    lget(&mut chunks[current], value_slot, line);
    emit_import_call(chunks, current, "ecma:array", "length", 1, line);
    let arr_len_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_len_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    // REF_IS_NULL returns I32(0/1) — use directly, same reason as above.
    chunks[current].emit_if_value(line);

    lget(&mut chunks[current], value_slot, line);
    emit_import_call(chunks, current, "ecma:map", "size", 1, line);
    let map_len_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_len_slot, line);

    // map_len != 0: ecma:map.size returns I32 — use I32_NE directly.
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_len_slot, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value_slot, line);
    emit_import_call(chunks, current, "ecma:object", "keys", 1, line);
    emit_import_call(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_len_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit(chunks: &mut [Chunk], current: usize, line: u32) {
    let from = chunks[current].add_import("ecma:array", "from");
    chunks[current].emit_call(from, 1, line);
    emit_len(chunks, current, line);
}
