#![allow(dead_code)]
use vybe_compiler::primitives::{class_slots, instructions::core_wasm};
use vybe_runtime::{Chunk, Op};
pub fn emit_add_changed(chunks: &mut [Chunk], current: usize, line: u32) {
    let set_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let existed_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_has(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, existed_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, existed_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_add(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_method_has(chunks: &mut [Chunk], current: usize, line: u32) {
    let dict_slot = chunks[current].local_count;
    let key_slot = dict_slot + 1;
    let has_slot = key_slot + 1;
    let keys_slot = has_slot + 1;
    chunks[current].alloc_scratch(4);
    let keys_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );

    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dict_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    fixture_is_nullish(chunks, current, line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    fixture_is_nullish(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    vybe_compiler::primitives::collections::emit_contains(chunks, current, line);
    let cast = chunks[current].add_import("wasm:js-boolean", "cast");
    chunks[current].emit_call(cast, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, has_slot, line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_bool_const(true, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    fixture_is_nullish(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn build_setdefault(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_setdefault");
    c.arity = 3;
    c.local_count = 5; // dict(0), key(1), default(2), existing(3), result(4)
    let dict = 0u16;
    let key = 1u16;
    let default = 2u16;
    let existing = 3u16;
    let result = 4u16;

    // existing = dict[key]
    c.emit_op_u16(Op::LOCAL_GET, dict, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, existing, 0);

    // result = existing (default to existing; overwrite if missing)
    c.emit_op_u16(Op::LOCAL_GET, existing, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // if existing is null/undefined: assign default + use it as result.
    let done_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, existing, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0); // existing not null → keep result, exit

    // dict[key] = default; result = default
    c.emit_op_u16(Op::LOCAL_GET, dict, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op_u16(Op::LOCAL_GET, default, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    c.emit_op_u16(Op::LOCAL_GET, default, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_end(0);
    c.patch_block(done_block);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn emit_has(c: &mut [Chunk], cur: usize, line: u32) {
    vybe_compiler::primitives::sets::emit_has(c, cur, line);
}
fn emit_add(c: &mut [Chunk], cur: usize, line: u32) {
    vybe_compiler::primitives::sets::emit_add(c, cur, line);
}
fn fixture_is_nullish(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let idx = chunks[current].add_import("wasm:js-undefined", "test");
    chunks[current].emit_call(idx, 1, line);
    chunks[current].emit_op(Op::I32_OR, line);
}
