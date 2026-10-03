//! Insertion-sort emission retained only as a benchmark reference.
//! Captured after scalar counter specialization, before the merge-sort change.
//! Keep the algorithm fixed; shared host/operation implementations are common
//! to the reference and production sorter for a conservative comparison.
use vybe_runtime::{Chunk, Op};

pub fn build_sorted(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sorted");
    c.arity = 1;
    c.local_count = 6; // arr(0) + result(1) + i(2) + j(3) + len(4) + key(5)
    let arr = 0u16;
    let result = 1;
    let i = 2;
    let j = 3;
    let len = 4;
    let key = 5;

    // Copy input array → result (so we don't mutate the original)
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_i32_const(i32::MAX, 0);
    vybe_compiler::primitives::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // len = result.length
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    vybe_compiler::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    // Insertion sort: for i = 1 to len-1
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    // Private counter and collection length are proven numeric producers.
    c.emit_op(Op::F64_LT, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = result[i]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);

    // while j >= 0 && result[j] > key
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    // j is a private i32 counter, initialized and decremented with i32 ops.
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(1, 0); // exit inner loop

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    vybe_compiler::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // result[j+1] = result[j]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    // Now stack: [result, j+1] — need value = result[j]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    vybe_compiler::primitives::collections::emit_get_into(imports, &mut c, 0);
    vybe_compiler::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0);
    c.patch_loop(inner_loop_p);
    c.emit_end(0);
    c.patch_block(inner_block_p);

    // result[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    vybe_compiler::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0);
    c.patch_loop(outer_loop_p);
    c.emit_end(0);
    c.patch_block(outer_block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
