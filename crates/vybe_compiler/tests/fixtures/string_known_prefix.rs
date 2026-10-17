use vybe_runtime::{Chunk, Op};
pub fn build_to_string(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_tostring");
    c.arity = 1;
    c.local_count = 1;
    let val = 0u16;
    c.emit_string_const("", 0);
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
pub fn build_str_insert(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_str_insert");
    c.arity = 3;
    c.local_count = 3;
    let value = 2u16;

    // prefix = str[0:index]
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    {
        let idx = c.add_import("wasm:js-string", "substring");
        c.emit_call(idx, 3, 0);
    }

    // prefix + value (keeps current coercion behavior for non-string values)
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);

    // + suffix = str[index:]
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_i32_const(i32::MAX, 0);
    {
        let idx = c.add_import("wasm:js-string", "substring");
        c.emit_call(idx, 3, 0);
    }
    {
        let idx = c.add_import("wasm:js-string", "concat");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op(Op::RETURN, 0);
    c
}
pub fn build_pascal_str_insert(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_str_insert");
    c.arity = 3;
    c.local_count = 3;
    let value = 0u16;
    let target = 1u16;
    let index = 2u16;
    let max = c.add_constant(vybe_runtime::Value::I32(i32::MAX));

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    emit_str_substring(&mut c, 0);

    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, max, 0);
    emit_str_substring(&mut c, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
pub fn build_pascal_str_remove_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_str_remove_range");
    c.arity = 3;
    c.local_count = 3;
    let target = 0u16;
    let start = 1u16;
    let count = 2u16;
    let max = c.add_constant(vybe_runtime::Value::I32(i32::MAX));

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    emit_str_substring(&mut c, 0);

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, max, 0);
    emit_str_substring(&mut c, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
fn emit_str_substring(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-string", "substring");
    c.emit_call(idx, 3, line);
}
