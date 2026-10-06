// Captured packet assembly: private spills followed by host pushes.
use vybe_runtime::{Chunk, Op};
pub fn emit_from_stack(chunks: &mut [Chunk], current: usize, n: u16, line: u32) {
    let base = chunks[current].alloc_scratch(n.max(1));
    vybe_compiler::primitives::collections::emit_pack_n(chunks, current, n, base, line);
    vybe_compiler::primitives::multivalue::emit_tag(chunks, current, line);
}

fn load(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn save(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn emit_is_multi_value_slot(c: &mut Chunk, slot: u16, line: u32) {
    vybe_compiler::primitives::multivalue::emit_is_multi_value_slot(c, slot, line);
}
fn emit_tag(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::multivalue::emit_tag(chunks, current, line);
}
pub fn emit_as_multi_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let row = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], value, line);
    emit_is_multi_value_slot(&mut chunks[current], value, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_else(line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], row, line);
    load(&mut chunks[current], row, line);
    load(&mut chunks[current], value, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], row, line);
    emit_tag(chunks, current, line);
    chunks[current].emit_end(line);
}

/// Read the first value from a packet, or pass through a scalar.
/// Stack: `[value_or_packet] -> [first_or_value]`.
pub fn emit_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], value, line);
    emit_is_multi_value_slot(&mut chunks[current], value, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}
