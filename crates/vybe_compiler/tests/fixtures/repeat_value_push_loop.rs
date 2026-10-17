use vybe_runtime::{Chunk, Op};
fn emit_import_call(
    cs: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let i = cs[current].add_import(module, name);
    cs[current].emit_call(i, argc, line);
}
pub fn emit_repeat_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let count = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);

    let done = chunks[current].emit_block(line);
    let (again, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_import_call(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(again);
    chunks[current].emit_end(line);
    chunks[current].patch_block(done);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}
