// Frozen pre-optimization array helper bodies.
use vybe_runtime::Chunk;
pub fn build_array_remove_value(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_remove_value");
    c.arity = 2; // arr, value
    c.local_count = 3;
    let arr = 0u16;
    let value = 1;
    let idx = 2;

    // idx = indexOf(arr, value)
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, value, 0);
    let index_of = c.add_import("ecma:array", "indexOf");
    c.emit_call(index_of, 2, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, idx, 0);

    // if idx >= 0: splice + return true
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, idx, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    vybe_compiler::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    c.emit_if(0);

    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, idx, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1); // deleteCount = 1
    let splice = c.add_import("ecma:array", "splice");
    c.emit_call(splice, 3, 0);
    c.emit_op(vybe_runtime::opcode::Op::DROP, 0);
    c.emit_bool_const(true, 0);
    c.emit_op(vybe_runtime::opcode::Op::RETURN, 0);

    c.emit_end(0);
    c.emit_bool_const(false, 0);
    c.emit_op(vybe_runtime::opcode::Op::RETURN, 0);
    c
}

// ── array_insert_range(arr, index, src) → null ──────────────────────────
// Loop: for i in 0..src.length: splice(arr, index+i, 0, src[i])
pub fn build_array_insert_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_insert_range");
    c.arity = 3; // arr, index, src
    c.local_count = 5;
    let arr = 0u16;
    let index = 1;
    let src = 2;
    let i = 3;
    let src_len = 4;

    let len_import = c.add_import("ecma:array", "length");
    let get_import = c.add_import("ecma:array", "get");
    let splice_import = c.add_import("ecma:array", "splice");

    // src_len = length(src)
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, src, 0);
    c.emit_call(len_import, 1, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, src_len, 0);
    // i = 0
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, i, 0);

    let blk = c.emit_block(0);
    let (lp, _) = c.emit_loop_s(0);
    // if i >= src_len break
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, i, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, src_len, 0);
    vybe_compiler::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);
    // splice(arr, index+i, 0, src[i])
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, index, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, src, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, i, 0);
    c.emit_call(get_import, 2, 0);
    c.emit_call(splice_import, 4, 0);
    c.emit_op(vybe_runtime::opcode::Op::DROP, 0);
    // i++
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, i, 0);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(lp);
    c.emit_end(0);
    c.patch_block(blk);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(vybe_runtime::opcode::Op::RETURN, 0);
    c
}

// ── array_set_range(arr, index, src) → null ─────────────────────────────
// Loop: for i in 0..src.length: arr[index+i] = src[i]
pub fn build_array_set_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_set_range");
    c.arity = 3;
    c.local_count = 5;
    let arr = 0u16;
    let index = 1;
    let src = 2;
    let i = 3;
    let src_len = 4;

    let len_import = c.add_import("ecma:array", "length");
    let get_import = c.add_import("ecma:array", "get");
    let set_import = c.add_import("ecma:array", "set");

    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, src, 0);
    c.emit_call(len_import, 1, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, src_len, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, i, 0);

    let blk = c.emit_block(0);
    let (lp, _) = c.emit_loop_s(0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, i, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, src_len, 0);
    vybe_compiler::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);
    // set(arr, index+i, get(src, i))
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, index, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, src, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, i, 0);
    c.emit_call(get_import, 2, 0);
    c.emit_call(set_import, 3, 0);
    c.emit_op(vybe_runtime::opcode::Op::DROP, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    vybe_compiler::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, i, 0);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(lp);
    c.emit_end(0);
    c.patch_block(blk);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(vybe_runtime::opcode::Op::RETURN, 0);
    c
}
