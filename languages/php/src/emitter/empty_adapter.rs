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
    // The predicate already returns the PHP Boolean required by a value
    // consumer. Do not cast it to i32 only to immediately box it again.
    crate::emitter::array_adapter::emit_php_empty_from_slot(chunks, current, value_slot, line);
}

/// Emit the same empty predicate without materializing its Boolean value.
/// Only the typed dispatcher may select this for a condition request.
pub(super) fn emit_php_empty_condition(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_php_empty(chunks, current, 1, line);
    // The shared predicate returns a Boolean; condition consumers require i32.
    let chunk = &mut chunks[current];
    let cast = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(cast, 1, line);
}
