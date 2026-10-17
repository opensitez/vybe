use vybe_runtime::{Chunk, Op};
pub fn build_any_all(imports: &mut Chunk, name: &str, is_any: bool) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = 1;
    c.local_count = 3; // arr(0) + i(1) + len(2)
    let arr = 0u16;
    let i = 1;
    let len = 2;

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
    c.emit_br_if(1, 0); // exit loop → fell through

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    vybe_compiler::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    if is_any {
        // any: if truthy → return true
        c.emit_if(0);
        c.emit_bool_const(true, 0);
        c.emit_op(Op::RETURN, 0);
        c.emit_end(0);
    } else {
        // all: if falsy → return false
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_if(0);
        c.emit_bool_const(false, 0);
        c.emit_op(Op::RETURN, 0);
        c.emit_end(0);
    }

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    // Loop fell through: any → false, all → true
    if is_any {
        c.emit_bool_const(false, 0);
    } else {
        c.emit_bool_const(true, 0);
    }
    c.emit_op(Op::RETURN, 0);
    c
}
