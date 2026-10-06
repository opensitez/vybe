use vybe_compiler::primitives::{class_slots, instructions::core_wasm};
use vybe_runtime::{Chunk, opcode::Op};

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
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    vybe_compiler::primitives::collections::emit_contains(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, has_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, has_slot, line);
    chunks[current].emit_if(line);

    chunks[current].emit_bool_const(true, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_method_delete(chunks: &mut [Chunk], current: usize, line: u32) {
    let dict_slot = chunks[current].local_count;
    let key_slot = dict_slot + 1;
    let idx_slot = dict_slot + 2;
    chunks[current].alloc_scratch(3);
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
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    vybe_compiler::primitives::collections::emit_index_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    vybe_compiler::primitives::collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);

    chunks[current].emit_else(line);

    chunks[current].emit_bool_const(false, line);

    chunks[current].emit_end(line);
}
