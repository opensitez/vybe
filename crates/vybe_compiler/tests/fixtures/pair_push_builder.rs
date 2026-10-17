use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::{Chunk, Op};
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
fn emit_import_call_into(
    _: &mut Chunk,
    code: &mut Chunk,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = code.add_import(module, name);
    code.emit_call(idx, argc, line);
}
pub fn emit_array_pair(chunks: &mut [Chunk], current: usize, line: u32) {
    let v2 = chunks[current].local_count;
    let v1 = chunks[current].local_count + 1;
    chunks[current].alloc_scratch(2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    emit_import_call(chunks, current, "ecma:array", "newWithLength", 1, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v1, line);
    emit_import_call(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v2, line);
    emit_import_call(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

pub fn emit_array_pair_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let v2 = code.local_count;
    let v1 = code.local_count + 1;
    code.alloc_scratch(2);
    // Stack: [v1, v2] — stash both into temp slots (peek-set + drop).
    code.emit_op_u16(Op::LOCAL_SET, v2, line);
    code.emit_op_u16(Op::LOCAL_SET, v1, line);
    // arr = ecma:array.newWithLength(0)
    core_wasm::i32_const(code, line, 0);
    emit_import_call_into(imports, code, "ecma:array", "newWithLength", 1, line);
    // arr.push(v1)
    code.emit_dup(line);
    code.emit_op_u16(Op::LOCAL_GET, v1, line);
    emit_import_call_into(imports, code, "ecma:array", "push", 2, line);
    code.emit_op(Op::DROP, line);
    // arr.push(v2)
    code.emit_dup(line);
    code.emit_op_u16(Op::LOCAL_GET, v2, line);
    emit_import_call_into(imports, code, "ecma:array", "push", 2, line);
    code.emit_op(Op::DROP, line);
}

pub fn build_enumerate(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_enumerate");
    c.arity = 1;
    c.local_count = 4; // arr(0) + result(1) + i(2) + len(3)
    let arr = 0u16;
    let result = 1;
    let i = 2;
    let len = 3;

    vybe_compiler::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    vybe_compiler::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    // Private counter and collection length are proven numeric producers.
    c.emit_op(Op::F64_LT, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(1, 0); // exit loop

    // Build pair [i, arr[i]], then push onto result.
    // array_push takes [array, value] — so emit result first, then pair.
    c.emit_op_u16(Op::LOCAL_GET, result, 0); // result on stack
    c.emit_op_u16(Op::LOCAL_GET, i, 0); // i
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0); // arr[i]
    emit_array_pair_into(imports, &mut c, 0); // pair = [i, arr[i]]
    vybe_compiler::primitives::collections::emit_push_into(imports, &mut c, 0); // result.push(pair)
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
