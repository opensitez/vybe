use vybe_runtime::{Chunk,Op};
use vybe_compiler::primitives::{class_slots,instructions::core_wasm};
use vybe_compiler::primitives::tuples::emit_tag;
pub fn emit_tuple(chunks: &mut [Chunk], current: usize, n: u16, line: u32) {
    let base = chunks[current].alloc_scratch(n.max(1));
    emit_pack_n(chunks, current, n, base, line);
    emit_tag(chunks, current, line);
    emit_positional_item_fields(chunks, current, n as usize, line);
}

fn emit_positional_item_fields(chunks: &mut [Chunk], current: usize, n: usize, line: u32) {
    for i in 0..n {
        emit_positional_item_field(chunks, current, i, &format!("Item{}", i + 1), line);
        emit_positional_item_field(chunks, current, i, &format!("item{}", i + 1), line);
    }
}

fn emit_positional_item_field(
    chunks: &mut [Chunk],
    current: usize,
    index: usize,
    name: &str,
    line: u32,
) {
    let c = &mut chunks[current];
    c.emit_dup(line);
    c.emit_dup(line);
    core_wasm::i32_const(c, line, index as i32);
    c.emit_op(Op::ARRAY_GET, line);
    let k = class_slots::resolve_interned(
        c,
        &class_slots::ClassSlot::internal(name),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        c,
        class_slots::ObjSource::Stack,
        &k,
        class_slots::ValueSource::Stack,
        line,
    );
}


pub fn emit_pack_n(chunks: &mut [Chunk], current: usize, n: u16, slot_base: u16, line: u32) {
    if n == 0 {
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    // Stash in reverse (stack top = v(N-1) goes into slot_base + N-1).
    for i in (0..n).rev() {
        let slot = slot_base + i;
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    // Build empty array, push each in forward order.
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    for i in 0..n {
        chunks[current].emit_dup(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot_base + i, line);
        { let idx=chunks[current].add_import("ecma:array", "push"); chunks[current].emit_call(idx,2,line); }
        chunks[current].emit_op(Op::DROP, line); // drop new_length
    }
}

