//! PHP `empty()` adapter.
//!
//! Kept as a focused `php.*` resolver leaf rather than living in the
//! miscellaneous adapter bucket.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

pub fn emit_php_empty(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);

    lset(chunk, value_slot, line);
    crate::emitter::array_adapter::emit_php_empty_from_slot(chunks, current, value_slot, line);
    // Comparison primitives produce an i32 condition. PHP exposes a bool,
    // including to strict comparisons and var_dump.
    let chunk = &mut chunks[current];
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}
