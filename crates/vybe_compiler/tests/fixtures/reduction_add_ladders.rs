use vybe_runtime::{Chunk, Op};
fn call1(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 1, line);
}
fn call2(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 2, line);
}
fn alloc_locals(chunk: &mut Chunk, n: u16) -> u16 {
    chunk.alloc_scratch(n)
}
fn save(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn load(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn f64_const(chunk: &mut Chunk, v: f64, line: u32) {
    chunk.emit_f64_const(v, line);
}
fn emit_bigint_binop(chunk: &mut Chunk, name: &str, line: u32) {
    let op = chunk.add_import("ecma:bigint", name);
    call2(chunk, op, line); // [a, b] -> BigInt
}
fn emit_js_to_number_f64(
    chunk: &mut Chunk,
    slot: u16,
    to_f64: u16,
    test_bool: u16,
    cast_bool: u16,
    line: u32,
) {
    load(chunk, slot, line);
    {
        let idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_if_value(line);
    f64_const(chunk, f64::NAN, line);
    chunk.emit_else(line);

    load(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    f64_const(chunk, 0.0, line);
    chunk.emit_else(line);

    load(chunk, slot, line);
    call1(chunk, test_bool, line);
    chunk.emit_if_value(line);
    load(chunk, slot, line);
    call1(chunk, cast_bool, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
    chunk.emit_else(line);
    load(chunk, slot, line);
    call1(chunk, to_f64, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
}
pub fn emit_dyn_add_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let slots = alloc_locals(code, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_str = code.add_import("wasm:js-string", "test");
    let str_cast = code.add_import("wasm:js-string", "cast");
    let str_concat = code.add_import("wasm:js-string", "concat");
    let str_from_f64 = code.add_import("wasm:js-string", "fromF64");
    let test_bigint = code.add_import("wasm:js-bigint", "test");
    let to_f64 = code.add_import("wasm:js-number", "toF64");
    let from_f64 = code.add_import("wasm:js-number", "fromF64");
    let test_bool = code.add_import("wasm:js-boolean", "test");
    let cast_bool = code.add_import("wasm:js-boolean", "cast");

    save(code, b_slot, line);
    save(code, a_slot, line);

    load(code, a_slot, line);
    call1(code, test_str, line);
    load(code, b_slot, line);
    call1(code, test_str, line);
    code.emit_op(Op::I32_OR, line);
    code.emit_if_value(line);
    load(code, a_slot, line);
    call1(code, test_str, line);
    code.emit_if_value(line);
    load(code, a_slot, line);
    call1(code, str_cast, line);
    code.emit_else(line);
    emit_js_to_number_f64(code, a_slot, to_f64, test_bool, cast_bool, line);
    call1(code, str_from_f64, line);
    code.emit_end(line);
    load(code, b_slot, line);
    call1(code, test_str, line);
    code.emit_if_value(line);
    load(code, b_slot, line);
    call1(code, str_cast, line);
    code.emit_else(line);
    emit_js_to_number_f64(code, b_slot, to_f64, test_bool, cast_bool, line);
    call1(code, str_from_f64, line);
    code.emit_end(line);
    call2(code, str_concat, line);

    code.emit_else(line);
    load(code, a_slot, line);
    call1(code, test_bigint, line);
    load(code, b_slot, line);
    call1(code, test_bigint, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_value(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    emit_bigint_binop(code, "add", line);
    code.emit_else(line);
    emit_js_to_number_f64(code, a_slot, to_f64, test_bool, cast_bool, line);
    emit_js_to_number_f64(code, b_slot, to_f64, test_bool, cast_bool, line);
    code.emit_op(Op::F64_ADD, line);
    call1(code, from_f64, line);
    code.emit_end(line);
    code.emit_end(line);
}

pub fn build_sum(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sum");
    c.arity = 1;
    c.local_count = 4; // arr(0) + total(1) + i(2) + len(3)
    let arr = 0u16;
    let total = 1;
    let i = 2;
    let len = 3;

    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, total, 0);

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

    c.emit_op_u16(Op::LOCAL_GET, total, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, total, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, total, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
