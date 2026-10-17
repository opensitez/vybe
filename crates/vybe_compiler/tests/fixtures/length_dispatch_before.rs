use vybe_compiler::primitives::collections::emit_import_call_into;
use vybe_runtime::{Chunk, Op};
pub fn emit_len_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    // Runtime-typed length implemented with structured control flow.
    //
    // WASM blocks with `() -> ()` type can't leave a value on the stack
    // at the end — to hand the computed length out of the dispatch we
    // stash it in a scratch local and reload after the block closes.
    //
    // Stack at entry: [val]. Exit: [len].
    let scratch_val = code.local_count;
    let scratch_len = code.local_count + 1;
    code.local_count = code
        .local_count
        .checked_add(2)
        .expect("emit_len_into: local slot overflow");
    code.emit_op_u16(Op::LOCAL_SET, scratch_val, line);

    let outer = code.emit_block(line);
    let str_block = code.emit_block(line);
    code.emit_op_u16(Op::LOCAL_GET, scratch_val, line);
    {
        let idx = code.add_import("wasm:js-string", "test");
        code.emit_call(idx, 1, line);
    }
    // wasm:js-string.test already yields i32(0/1); invert with I32_EQZ. Do NOT use
    // emit_dyn_not here — it calls emit_dyn_to_bool, which registers imports
    // on `code` via add_import. Those collide with the chunks[0]-based global
    // import indices used by emit_import_call_into below, making CALL_IMPORT
    // resolve the wrong host fn (ecma:array.length → wasm:js-string.length).
    code.emit_op(Op::I32_EQZ, line);
    // `br_if 0` pops the i32. When `!is_string` is true we jump out of
    // `str_block`, falling through to the array-length branch.
    code.emit_br_if(0, line);
    // String path — stash length and exit outer block.
    code.emit_op_u16(Op::LOCAL_GET, scratch_val, line);
    emit_import_call_into(imports, code, "wasm:js-string", "length", 1, line);
    code.emit_op_u16(Op::LOCAL_SET, scratch_len, line);
    code.emit_br(1, line);
    code.emit_end(line);
    code.patch_block(str_block);
    // Array path — fallthrough from the `br_if 0`.
    code.emit_op_u16(Op::LOCAL_GET, scratch_val, line);
    emit_import_call_into(imports, code, "ecma:array", "length", 1, line);
    code.emit_op_u16(Op::LOCAL_SET, scratch_len, line);
    code.emit_end(line);
    code.patch_block(outer);
    code.emit_op_u16(Op::LOCAL_GET, scratch_len, line);
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
    emit_len_into(imports, &mut c, 0);
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
pub fn build_min(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_min");
    c.arity = 1;
    c.local_count = 4; // arr(0) + best(1) + i(2) + len(3)
    let arr = 0u16;
    let best = 1;
    let i = 2;
    let len = 3;

    // best = arr[0]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    // Private counter and collection length are proven numeric producers.
    c.emit_op(Op::F64_LT, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(1, 0); // exit loop

    // if arr[i] < best: best = arr[i]
    // block must wrap ALL condition operands + comparison + body
    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    vybe_compiler::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0); // skip if NOT less than
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_end(0);
    c.patch_block(skip_block_p);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
pub fn build_max(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_max");
    c.arity = 1;
    c.local_count = 4;
    let arr = 0u16;
    let best = 1;
    let i = 2;
    let len = 3;

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    // Private counter and collection length are proven numeric producers.
    c.emit_op(Op::F64_LT, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(1, 0); // exit loop

    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    vybe_compiler::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0); // skip if NOT greater than
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_end(0);
    c.patch_block(skip_block_p);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
pub fn build_isempty(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isempty");
    c.arity = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_len_into(imports, &mut c, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op(Op::F64_EQ, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
