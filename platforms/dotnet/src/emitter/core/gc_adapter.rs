use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;

pub fn emit_gc_noop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_gc_zero(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_i32_const(0, line);
}

pub fn emit_gc_get_total_memory(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_f64_const(0.0, line);
}

pub fn emit_gc_get_generation(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
}
