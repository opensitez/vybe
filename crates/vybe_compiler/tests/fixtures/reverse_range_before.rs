use vybe_runtime::Chunk;
pub fn build_array_reverse_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_reverse_range");
    c.arity = 3;
    c.local_count = 6;
    let arr = 0u16;
    let index = 1;
    let count = 2;
    let lo = 3;
    let hi = 4;
    let tmp = 5;

    let get_import = c.add_import("ecma:array", "get");
    let set_import = c.add_import("ecma:array", "set");

    // lo = index; hi = index + count - 1
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, index, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, lo, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, index, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, count, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    // The previous constant negation produces precisely boxed F64(-1).
    c.emit_f64_const(-1.0, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, hi, 0);

    let blk = c.emit_block(0);
    let (lp, _) = c.emit_loop_s(0);
    // while lo < hi
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, lo, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, hi, 0);
    vybe_compiler::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    c.emit_op(vybe_runtime::opcode::Op::I32_EQZ, 0);
    c.emit_br_if(1, 0);
    // tmp = arr[lo]
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, lo, 0);
    c.emit_call(get_import, 2, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, tmp, 0);
    // arr[lo] = arr[hi]
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, lo, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, hi, 0);
    c.emit_call(get_import, 2, 0);
    c.emit_call(set_import, 3, 0);
    c.emit_op(vybe_runtime::opcode::Op::DROP, 0);
    // arr[hi] = tmp
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, hi, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, tmp, 0);
    c.emit_call(set_import, 3, 0);
    c.emit_op(vybe_runtime::opcode::Op::DROP, 0);
    // lo++; hi--
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, lo, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, lo, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, hi, 0);
    // The previous constant negation produces precisely boxed F64(-1).
    c.emit_f64_const(-1.0, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, hi, 0);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(lp);
    c.emit_end(0);
    c.patch_block(blk);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(vybe_runtime::opcode::Op::RETURN, 0);
    c
}
