//! Java `String.join` — the ONE emitter this crate still owns, because
//! the rest of `java.lang.String` now lives in `platforms/jvm`.

use vybe_compiler::primitives::collections;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_join(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let elem_count = argc.saturating_sub(1);
    let first_elem_slot = chunks[current].alloc_scratch(elem_count as u16);
    let delim_slot = chunks[current].alloc_scratch(1);

    for k in (0..elem_count).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, first_elem_slot + k as u16, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, delim_slot, line);

    if elem_count == 1 {
        let elem_slot = first_elem_slot;
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
        let len = chunks[current].add_import("ecma:array", "length");
        chunks[current].emit_call(len, 1, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        collections::emit_array_new(chunks, current, 0, line);
        let array_slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
        let push_idx = chunks[current].add_import("ecma:array", "push");
        chunks[current].emit_call(push_idx, 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, delim_slot, line);
        let join_idx = chunks[current].add_import("ecma:array", "join");
        chunks[current].emit_call(join_idx, 2, line);
        return;
    }

    collections::emit_array_new(chunks, current, 0, line);
    let array_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    let push_idx = chunks[current].add_import("ecma:array", "push");
    for k in 0..elem_count {
        chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, first_elem_slot + k as u16, line);
        chunks[current].emit_call(push_idx, 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delim_slot, line);
    let join_idx = chunks[current].add_import("ecma:array", "join");
    chunks[current].emit_call(join_idx, 2, line);
}

