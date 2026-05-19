use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

pub(super) fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].local_count;
    chunks[current].local_count += argc as u16;
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    base
}

pub(super) fn emit_null_result(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op(Op::NULL, line);
}