use vybe_runtime::{Chunk, Op};
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
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
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
