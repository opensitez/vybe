use vybe_runtime::{Chunk, Op};
pub fn emit_clamp(chunk: &mut Chunk, line: u32) {
    let max_slot = chunk.local_count;
    let min_slot = chunk.local_count + 1;
    chunk.alloc_scratch(2);
    chunk.emit_op_u16(Op::LOCAL_SET, max_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, min_slot, line);
    // stack: [x]
    chunk.emit_op_u16(Op::LOCAL_GET, min_slot, line); // [x, min]
    chunk.emit_op(Op::F64_MAX, line); // [max(x, min)]
    chunk.emit_op_u16(Op::LOCAL_GET, max_slot, line); // [max(x,min), max]
    chunk.emit_op(Op::F64_MIN, line); // [min(max(x,min), max)]
}

pub fn emit_signbit(chunk: &mut Chunk, line: u32) {
    let base = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, base, line); // stash x
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op(Op::F64_COPYSIGN, line); // ±1
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line); // < 0 → i32
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}
