//! Previous append-and-sort insertion retained as a fixed benchmark reference.
use vybe_compiler::primitives::collections;
use vybe_runtime::{Chunk, Op};
pub fn emit_push_with_comparator(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let cmp = chunk.alloc_scratch(1);
    let x = chunk.alloc_scratch(1);
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, cmp, line);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);

    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cmp, line);
    let _ = chunk;
    collections::emit_sort_with_comparator(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
