// Frozen dictionary helper factories before i32 predicate pass.
use vybe_runtime::{Chunk, Op};
pub fn build_dict_values_from_entries(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_dict_values_from_entries");
    c.arity = 1;
    c.local_count = 4; // entries(0), result(1), i(2), len(3)
    let entries = 0u16;
    let result = 1u16;
    let i = 2u16;
    let len = 3u16;

    // result = []
    vybe_compiler::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // len = entries.length
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    vybe_compiler::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    // i = 0
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // if i >= len: break
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    vybe_compiler::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    c.emit_op(vybe_runtime::opcode::Op::I32_EQZ, 0);
    c.emit_br_if(1, 0);

    // result.push(entries[i][1])
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::ARRAY_GET, 0);
    vybe_compiler::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

pub fn build_setdefault(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_setdefault");
    c.arity = 3;
    c.local_count = 5; // dict(0), key(1), default(2), existing(3), result(4)
    let dict = 0u16;
    let key = 1u16;
    let default = 2u16;
    let existing = 3u16;
    let result = 4u16;

    // existing = dict[key]
    c.emit_op_u16(Op::LOCAL_GET, dict, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, existing, 0);

    // result = existing (default to existing; overwrite if missing)
    c.emit_op_u16(Op::LOCAL_GET, existing, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // if existing is null/undefined: assign default + use it as result.
    let done_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, existing, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    vybe_compiler::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // existing not null → keep result, exit

    // dict[key] = default; result = default
    c.emit_op_u16(Op::LOCAL_GET, dict, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op_u16(Op::LOCAL_GET, default, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    c.emit_op_u16(Op::LOCAL_GET, default, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_end(0);
    c.patch_block(done_block);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
