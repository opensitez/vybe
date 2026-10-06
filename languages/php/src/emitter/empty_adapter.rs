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
    emit_php_empty_condition(chunks, current, line);
    // Value positions must retain PHP's Boolean representation.
    let chunk = &mut chunks[current];
    let from_i32 = chunk.add_import("wasm:js-boolean", "fromI32");
    chunk.emit_call(from_i32, 1, line);
}

/// Emit the same empty predicate without materializing its Boolean value.
/// Only the typed dispatcher may select this for a condition request.
pub(super) fn emit_php_empty_condition(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);

    lset(chunk, value_slot, line);
    crate::emitter::array_adapter::emit_php_empty_from_slot(chunks, current, value_slot, line);
}
