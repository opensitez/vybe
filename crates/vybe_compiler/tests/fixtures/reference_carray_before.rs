use vybe_compiler::primitives::class_slots;

use vybe_compiler::primitives::instructions::recipes;
use vybe_compiler::primitives::pointers::{
    CARRAY_BASE_KEY, CARRAY_IDX_KEY, CARRAY_KIND, CELL_KIND, REF_KIND_KEY, REF_VALUE_KEY,
    SHARED_ADDR_KEY, SHARED_KIND,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::references::emit_cell_load;
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn emit_field_get(chunks: &mut [Chunk], current: usize, field: &str, line: u32) {
    let key = class_slots::resolve(
        &class_slots::ClassSlot::internal(field),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &key,
        class_slots::Dest::Stack,
        line,
    );
}

fn emit_kind_eq(chunks: &mut [Chunk], current: usize, obj_slot: u16, kind_key: &str, line: u32) {
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, REF_KIND_KEY, line);
    let kind_slot = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], kind_slot, line);
    vybe_compiler::primitives::ops::emit_string_slot_eq_literal(
        &mut chunks[current],
        kind_slot,
        kind_key,
        line,
    );
}

pub fn build_autoderef() -> Chunk {
    let mut chunk = Chunk::new("__stdlib_autoderef");
    chunk.arity = 1;
    chunk.local_count = 1;
    lget(&mut chunk, 0, 0);
    let mut chunks = vec![chunk];
    emit_autoderef_inline(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    chunks.pop().unwrap()
}

fn emit_autoderef_inline(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj_slot = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], obj_slot, line);

    lget(&mut chunks[current], obj_slot, line);
    recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    emit_autoderef_object_inline(chunks, current, obj_slot, line);

    chunks[current].emit_else(line);
    lget(&mut chunks[current], obj_slot, line);
    chunks[current].emit_end(line);
}

fn emit_autoderef_object_inline(chunks: &mut [Chunk], current: usize, obj_slot: u16, line: u32) {
    // Ordinary objects have no reference tag. Return them without walking
    // the cell/carray/shared dispatch ladder at every variable read.
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, REF_KIND_KEY, line);
    let kind_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, kind_slot, line);
    let test = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_call(test, 1, line);
    chunks[current].emit_if_value(line);

    // The wrapper's tag is stable throughout this read. Once its string type
    // is established, dispatch on that value without repeating field access
    // and type testing for each reference representation.
    let equals = chunks[current].add_import("wasm:js-string", "equals");

    lget(&mut chunks[current], kind_slot, line);
    chunks[current].emit_string_const(CELL_KIND, line);
    chunks[current].emit_call(equals, 2, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], obj_slot, line);
    emit_cell_load(chunks, current, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], kind_slot, line);
    chunks[current].emit_string_const(CARRAY_KIND, line);
    chunks[current].emit_call(equals, 2, line);
    chunks[current].emit_if_value(line);

    let base_slot = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, CARRAY_BASE_KEY, line);
    lset(&mut chunks[current], base_slot, line);

    lget(&mut chunks[current], base_slot, line);
    recipes::is_object(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    emit_kind_eq(chunks, current, base_slot, CELL_KIND, line);
    chunks[current].emit_if_params(line, 0, 1);
    lget(&mut chunks[current], base_slot, line);
    emit_cell_load(chunks, current, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], base_slot, line);
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, CARRAY_IDX_KEY, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    lget(&mut chunks[current], base_slot, line);
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, CARRAY_IDX_KEY, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);

    lget(&mut chunks[current], kind_slot, line);
    chunks[current].emit_string_const(SHARED_KIND, line);
    chunks[current].emit_call(equals, 2, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], obj_slot, line);
    emit_field_get(chunks, current, SHARED_ADDR_KEY, line);
    vybe_compiler::primitives::threading::emit_atomic_load(&mut chunks[current], line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], obj_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    lget(&mut chunks[current], obj_slot, line);
    chunks[current].emit_end(line);
}
