use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_emitter::instructions::core_wasm;

use vybe_emitter::collections;

const VB_COLLECTION_ITEMS: &str = "__dotnet_vb_collection_items";
const VB_COLLECTION_KEYS: &str = "__dotnet_vb_collection_keys";
const BLOCKING_ITEMS: &str = "__dotnet_blocking_items";
const BLOCKING_CAPACITY: &str = "__dotnet_blocking_capacity";
const BLOCKING_COMPLETED: &str = "__dotnet_blocking_completed";
const BLOCKING_LIFO: &str = "__dotnet_blocking_lifo";
const OBSERVABLE_ITEMS: &str = "__dotnet_observable_items";
const OBSERVABLE_NOTIFYING: &str = "__dotnet_observable_notifying";

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    // Register on the chunk that EMITS the call, not chunk[0]. The VM's
    // `resolve_chunk_import` checks the executing chunk's own import table
    // first, so a chunk[0] (global) index baked into a function chunk resolves
    // to a wrong LOCAL import whenever it falls within that chunk's table —
    // which is exactly what happened once the .NET prelude stopped padding
    // chunk[0]'s import prefix. A local index always resolves correctly.
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

fn emit_string_const(chunks: &mut [Chunk], current: usize, value: &str, line: u32) {
    let idx = chunks[current].add_constant(Value::String(Arc::from(value)));
    chunks[current].emit_op_u16(Op::CONST, idx, line);
}

fn emit_get_field(chunks: &mut [Chunk], current: usize, object_slot: u16, field: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object_slot, line);
    emit_string_const(chunks, current, field, line);
    collections::emit_get(chunks, current, line);
}

fn emit_set_field(
    chunks: &mut [Chunk],
    current: usize,
    object_slot: u16,
    field: &str,
    value_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object_slot, line);
    emit_string_const(chunks, current, field, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_array_with_slots(chunks: &mut [Chunk], current: usize, slots: &[u16], line: u32) {
    for slot in slots {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
    }
    collections::emit_array_new(chunks, current, slots.len() as u16, line);
}

fn emit_slot_is_nullish(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-undefined", "test", 1, line);
    chunks[current].emit_op(Op::I32_OR, line);
}

fn emit_observable_items_slot(chunks: &mut [Chunk], current: usize, recv: u16, line: u32) -> u16 {
    let items = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items, line);
    chunks[current].emit_else(line);
    emit_get_field(chunks, current, recv, OBSERVABLE_ITEMS, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items, line);
    emit_slot_is_nullish(chunks, current, items, line);
    chunks[current].emit_if(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items, line);
    emit_set_field(chunks, current, recv, OBSERVABLE_ITEMS, items, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    items
}

pub fn emit_set_new_ignore_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    call_import(chunks, current, "ecma:set", "new", 0, line);
}

pub fn emit_list_new_from_iterable(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_spread_iterable(chunks, current, line);
}

pub fn emit_readonly_observable_collection_new(_chunks: &mut [Chunk], _current: usize, _line: u32) {
    // ReadOnlyObservableCollection is a view over the source ObservableCollection:
    // preserve the backing object so Count/Item and CollectionChanged stay in sync.
}

pub fn emit_property_changed_event_args_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let property_name = stash_args(chunks, current, 1, line);
    let args = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, args, line);
    emit_set_field(chunks, current, args, "PropertyName", property_name, line);
    emit_set_field(chunks, current, args, "propertyname", property_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, args, line);
}

pub fn emit_notify_collection_changed_event_args_new(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
) {
    let action = stash_args(chunks, current, 1, line);
    let args = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, args, line);
    emit_set_field(chunks, current, args, "Action", action, line);
    emit_set_field(chunks, current, args, "action", action, line);

    let minus_one = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, minus_one, line);
    emit_set_field(chunks, current, args, "OldStartingIndex", minus_one, line);
    emit_set_field(chunks, current, args, "oldstartingindex", minus_one, line);
    emit_set_field(chunks, current, args, "NewStartingIndex", minus_one, line);
    emit_set_field(chunks, current, args, "newstartingindex", minus_one, line);

    let null_value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, null_value, line);
    emit_set_field(chunks, current, args, "OldItems", null_value, line);
    emit_set_field(chunks, current, args, "olditems", null_value, line);
    emit_set_field(chunks, current, args, "NewItems", null_value, line);
    emit_set_field(chunks, current, args, "newitems", null_value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, args, line);
}

pub fn emit_set_new_from_iterable(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_spread_iterable(chunks, current, line);
    call_import(chunks, current, "ecma:set", "new", 1, line);
}

pub fn emit_hashset_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let value = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:set", "has", 2, line);
    let present_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, present_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, present_slot, line);
    chunks[current].emit_if(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:set", "add", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_end(line);
}

pub fn emit_dict_new_ignore_arg(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    call_import(chunks, current, "ecma:map", "new", 0, line);
}

pub fn emit_dict_try_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let recv = base;
    let key = base + 1;
    let value = base + 2;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_end(line);
}

pub fn emit_dict_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let arg = base + 1;
    let key = chunks[current].alloc_scratch(4);
    let value = key + 1;
    let current_value = key + 2;
    let len = key + 3;

    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, current_value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, current_value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    call_import(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_end(line);
}

pub fn emit_dict_ensure_capacity(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
}

pub fn emit_dict_trim_excess(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_blocking_field(
    chunks: &mut [Chunk],
    current: usize,
    recv_slot: u16,
    field: &str,
    line: u32,
) {
    emit_get_field(chunks, current, recv_slot, field, line);
}

pub fn emit_blocking_collection_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let input_slot = chunks[current].alloc_scratch(1);
    if argc > 0 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, input_slot, line);
    } else {
        chunks[current].emit_i32_const(-1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, input_slot, line);
    }

    let obj_slot = chunks[current].alloc_scratch(4);
    let items_slot = obj_slot + 1;
    let capacity_slot = obj_slot + 2;
    let lifo_slot = obj_slot + 3;

    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    if argc > 0 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
        call_import(chunks, current, "wasm:js-number", "test", 1, line);
        chunks[current].emit_if(line);
        chunks[current].emit_i32_const(0, line);
        call_import(chunks, current, "ecma:array", "newWithLength", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, items_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, capacity_slot, line);
        core_wasm::bool_const(&mut chunks[current], line, false);
        chunks[current].emit_op_u16(Op::LOCAL_SET, lifo_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, items_slot, line);
        chunks[current].emit_i32_const(-1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, capacity_slot, line);
        core_wasm::bool_const(&mut chunks[current], line, true);
        chunks[current].emit_op_u16(Op::LOCAL_SET, lifo_slot, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_i32_const(0, line);
        call_import(chunks, current, "ecma:array", "newWithLength", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, items_slot, line);
        chunks[current].emit_i32_const(-1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, capacity_slot, line);
        core_wasm::bool_const(&mut chunks[current], line, false);
        chunks[current].emit_op_u16(Op::LOCAL_SET, lifo_slot, line);
    }

    emit_set_field(chunks, current, obj_slot, BLOCKING_ITEMS, items_slot, line);
    emit_set_field(
        chunks,
        current,
        obj_slot,
        BLOCKING_CAPACITY,
        capacity_slot,
        line,
    );
    let completed_slot = chunks[current].alloc_scratch(1);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, completed_slot, line);
    emit_set_field(
        chunks,
        current,
        obj_slot,
        BLOCKING_COMPLETED,
        completed_slot,
        line,
    );
    emit_set_field(chunks, current, obj_slot, BLOCKING_LIFO, lifo_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_blocking_collection_try_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let value = base + 1;
    let slots = chunks[current].alloc_scratch(3);
    let items = slots;
    let capacity = slots + 1;
    let len = slots + 2;

    emit_blocking_field(chunks, current, recv, BLOCKING_ITEMS, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items, line);
    emit_blocking_field(chunks, current, recv, BLOCKING_CAPACITY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, capacity, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, capacity, line);
    chunks[current].emit_i32_const(0, line);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, capacity, line);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_end(line);
    chunks[current].emit_if_value(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_end(line);
}

pub fn emit_blocking_collection_add(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_blocking_collection_try_add(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_blocking_collection_take(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        chunks[current].emit_op(Op::DROP, line);
    }
    let base = stash_args(chunks, current, 1, line);
    let recv = base;
    let items = chunks[current].alloc_scratch(2);
    let lifo = items + 1;

    emit_blocking_field(chunks, current, recv, BLOCKING_ITEMS, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items, line);
    emit_blocking_field(chunks, current, recv, BLOCKING_LIFO, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lifo, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lifo, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    collections::emit_pop(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    collections::emit_shift(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_blocking_collection_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    emit_blocking_field(chunks, current, recv, BLOCKING_ITEMS, line);
    collections::emit_len(chunks, current, line);
}

pub fn emit_blocking_collection_complete_adding(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    let value = chunks[current].alloc_scratch(1);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    emit_set_field(chunks, current, recv, BLOCKING_COMPLETED, value, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_blocking_collection_is_completed(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    emit_blocking_field(chunks, current, recv, BLOCKING_COMPLETED, line);
}

pub fn emit_blocking_collection_items(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    emit_blocking_field(chunks, current, recv, BLOCKING_ITEMS, line);
}

pub fn emit_observable_collection_items(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    let items = emit_observable_items_slot(chunks, current, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
}

pub fn emit_observable_collection_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    let items = emit_observable_items_slot(chunks, current, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    collections::emit_len(chunks, current, line);
}

fn emit_observable_event_args(
    chunks: &mut [Chunk],
    current: usize,
    action: &str,
    old_items: Option<u16>,
    new_items: Option<u16>,
    old_index: Option<u16>,
    new_index: Option<u16>,
    line: u32,
) -> u16 {
    let args = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, args, line);

    let action_value = chunks[current].alloc_scratch(1);
    emit_string_const(chunks, current, action, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, action_value, line);
    emit_set_field(chunks, current, args, "Action", action_value, line);
    emit_set_field(chunks, current, args, "action", action_value, line);

    let old_index_value = chunks[current].alloc_scratch(1);
    if let Some(slot) = old_index {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunks[current].emit_i32_const(-1, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, old_index_value, line);
    emit_set_field(
        chunks,
        current,
        args,
        "OldStartingIndex",
        old_index_value,
        line,
    );
    emit_set_field(
        chunks,
        current,
        args,
        "oldstartingindex",
        old_index_value,
        line,
    );

    let new_index_value = chunks[current].alloc_scratch(1);
    if let Some(slot) = new_index {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunks[current].emit_i32_const(-1, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, new_index_value, line);
    emit_set_field(
        chunks,
        current,
        args,
        "NewStartingIndex",
        new_index_value,
        line,
    );
    emit_set_field(
        chunks,
        current,
        args,
        "newstartingindex",
        new_index_value,
        line,
    );

    let old_items_value = chunks[current].alloc_scratch(1);
    if let Some(slot) = old_items {
        emit_array_with_slots(chunks, current, &[slot], line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, old_items_value, line);
    emit_set_field(chunks, current, args, "OldItems", old_items_value, line);
    emit_set_field(chunks, current, args, "olditems", old_items_value, line);

    let new_items_value = chunks[current].alloc_scratch(1);
    if let Some(slot) = new_items {
        emit_array_with_slots(chunks, current, &[slot], line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, new_items_value, line);
    emit_set_field(chunks, current, args, "NewItems", new_items_value, line);
    emit_set_field(chunks, current, args, "newitems", new_items_value, line);

    args
}

fn emit_property_changed_args(
    chunks: &mut [Chunk],
    current: usize,
    property_name: &str,
    line: u32,
) -> u16 {
    let args = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, args, line);

    let value = chunks[current].alloc_scratch(1);
    emit_string_const(chunks, current, property_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    emit_set_field(chunks, current, args, "PropertyName", value, line);
    emit_set_field(chunks, current, args, "propertyname", value, line);
    args
}

fn emit_delegate_event(
    chunks: &mut [Chunk],
    current: usize,
    recv: u16,
    event_name: &str,
    fallback_name: &str,
    args: u16,
    line: u32,
) {
    let delegate = chunks[current].alloc_scratch(1);
    emit_get_field(chunks, current, recv, event_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, delegate, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_get_field(chunks, current, recv, fallback_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, delegate, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, args, line);
    vybe_emitter::delegates::emit_invoke(chunks, current, 3, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_throw_dotnet_exception(
    chunks: &mut [Chunk],
    current: usize,
    exception_name: &str,
    message: &str,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const(message, line);
    vybe_emitter::errors::emit_exception_new_finalize(&mut chunks[current], exception_name, line);
    vybe_emitter::errors::emit_throw(&mut chunks[current], line);
}

fn emit_check_index_in_range(
    chunks: &mut [Chunk],
    current: usize,
    recv: u16,
    index: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(0, line);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);

    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "ArgumentOutOfRangeException",
        "Index was out of range. Must be non-negative and less than the size of the collection.",
        line,
    );
    chunks[current].emit_end(line);
}

fn emit_check_not_reentrant(chunks: &mut [Chunk], current: usize, recv: u16, line: u32) {
    emit_get_field(chunks, current, recv, OBSERVABLE_NOTIFYING, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "InvalidOperationException",
        "Cannot change ObservableCollection during a CollectionChanged event.",
        line,
    );
    chunks[current].emit_end(line);
}

fn emit_set_observable_notifying(
    chunks: &mut [Chunk],
    current: usize,
    recv: u16,
    notifying: bool,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    core_wasm::bool_const(&mut chunks[current], line, notifying);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    emit_set_field(chunks, current, recv, OBSERVABLE_NOTIFYING, value, line);
}

fn emit_observable_collection_changed(
    chunks: &mut [Chunk],
    current: usize,
    recv: u16,
    args: u16,
    line: u32,
) {
    emit_set_observable_notifying(chunks, current, recv, true, line);
    emit_delegate_event(
        chunks,
        current,
        recv,
        "CollectionChanged",
        "collectionchanged",
        args,
        line,
    );
    emit_set_observable_notifying(chunks, current, recv, false, line);
}

fn emit_observable_property_changed(
    chunks: &mut [Chunk],
    current: usize,
    recv: u16,
    property_name: &str,
    line: u32,
) {
    let args = emit_property_changed_args(chunks, current, property_name, line);
    emit_delegate_event(
        chunks,
        current,
        recv,
        "PropertyChanged",
        "propertychanged",
        args,
        line,
    );
}

fn emit_observable_count_and_indexer_changed(
    chunks: &mut [Chunk],
    current: usize,
    recv: u16,
    line: u32,
) {
    emit_observable_property_changed(chunks, current, recv, "Count", line);
    emit_observable_property_changed(chunks, current, recv, "Item[]", line);
}

pub fn emit_observable_collection_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let value = base + 1;
    let items = emit_observable_items_slot(chunks, current, recv, line);

    emit_check_not_reentrant(chunks, current, recv, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    let index = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);

    emit_observable_count_and_indexer_changed(chunks, current, recv, line);
    let args = emit_observable_event_args(
        chunks,
        current,
        "Add",
        None,
        Some(value),
        None,
        Some(index),
        line,
    );
    emit_observable_collection_changed(chunks, current, recv, args, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_observable_collection_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let value = base + 1;
    let items = emit_observable_items_slot(chunks, current, recv, line);
    let index = chunks[current].alloc_scratch(1);

    emit_check_not_reentrant(chunks, current, recv, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_index_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(0, line);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_observable_count_and_indexer_changed(chunks, current, recv, line);
    let args = emit_observable_event_args(
        chunks,
        current,
        "Remove",
        Some(value),
        None,
        Some(index),
        None,
        line,
    );
    emit_observable_collection_changed(chunks, current, recv, args, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_end(line);
}

pub fn emit_observable_collection_remove_at(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let index = base + 1;
    let items = emit_observable_items_slot(chunks, current, recv, line);
    let value = chunks[current].alloc_scratch(1);

    emit_check_not_reentrant(chunks, current, recv, line);
    emit_check_index_in_range(chunks, current, items, index, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_observable_count_and_indexer_changed(chunks, current, recv, line);
    let args = emit_observable_event_args(
        chunks,
        current,
        "Remove",
        Some(value),
        None,
        Some(index),
        None,
        line,
    );
    emit_observable_collection_changed(chunks, current, recv, args, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_observable_collection_insert(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let recv = base;
    let index = base + 1;
    let value = base + 2;
    let items = emit_observable_items_slot(chunks, current, recv, line);

    emit_check_not_reentrant(chunks, current, recv, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_insert(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_observable_count_and_indexer_changed(chunks, current, recv, line);
    let args = emit_observable_event_args(
        chunks,
        current,
        "Add",
        None,
        Some(value),
        None,
        Some(index),
        line,
    );
    emit_observable_collection_changed(chunks, current, recv, args, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_observable_collection_set_index(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let recv = base;
    let index = base + 1;
    let value = base + 2;
    let items = emit_observable_items_slot(chunks, current, recv, line);
    let old_value = chunks[current].alloc_scratch(1);

    emit_check_not_reentrant(chunks, current, recv, line);
    emit_check_index_in_range(chunks, current, items, index, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, old_value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_observable_property_changed(chunks, current, recv, "Item[]", line);
    let args = emit_observable_event_args(
        chunks,
        current,
        "Replace",
        Some(old_value),
        Some(value),
        Some(index),
        Some(index),
        line,
    );
    emit_observable_collection_changed(chunks, current, recv, args, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_observable_collection_move(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let recv = base;
    let old_index = base + 1;
    let new_index = base + 2;
    let items = emit_observable_items_slot(chunks, current, recv, line);
    let value = chunks[current].alloc_scratch(1);

    emit_check_not_reentrant(chunks, current, recv, line);
    emit_check_index_in_range(chunks, current, items, old_index, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old_index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old_index, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, new_index, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_insert(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_observable_property_changed(chunks, current, recv, "Item[]", line);
    let args = emit_observable_event_args(
        chunks,
        current,
        "Move",
        Some(value),
        Some(value),
        Some(old_index),
        Some(new_index),
        line,
    );
    emit_observable_collection_changed(chunks, current, recv, args, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_observable_collection_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    let items = emit_observable_items_slot(chunks, current, recv, line);
    emit_check_not_reentrant(chunks, current, recv, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    collections::emit_clear(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_observable_count_and_indexer_changed(chunks, current, recv, line);
    let args = emit_observable_event_args(chunks, current, "Reset", None, None, None, None, line);
    emit_observable_collection_changed(chunks, current, recv, args, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_observable_collection_on_changed(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let args = base + 1;
    emit_observable_collection_changed(chunks, current, recv, args, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// Normalize the `IEnumerable<T>` argument in `slot` to an ECMA Set: drain it
/// to an array via the shared iterator protocol (handles arrays, `List<T>`,
/// other sets, and generators alike — `ecma:set.new` alone mishandles a Set
/// argument), then build a Set from that array. Stores the Set back in `slot`.
fn normalize_arg_to_set(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    call_import(chunks, current, "ecma:set", "new", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn emit_hashset_mutation(chunks: &mut [Chunk], current: usize, func: &str, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    let result_slot = chunks[current].alloc_scratch(4);
    let arr_slot = result_slot + 1;
    let idx_slot = result_slot + 2;
    let value_slot = result_slot + 3;

    // C# set methods take any `IEnumerable<T>`; the ECMA set operations require
    // a set-like operand. Normalize the argument to an ECMA Set via
    // `new Set(iterable)` (§24.2.1.1) so arrays / lists / generators all work —
    // pure adaptation, no host or VM change.
    normalize_arg_to_set(chunks, current, src, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    call_import(chunks, current, "ecma:set", func, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:set", "clear", 1, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_import(chunks, current, "ecma:set", "add", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_hashset_union_with(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    let arr_slot = chunks[current].alloc_scratch(3);
    let idx_slot = arr_slot + 1;
    let value_slot = arr_slot + 2;

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_import(chunks, current, "ecma:set", "add", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_hashset_intersect_with(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_hashset_mutation(chunks, current, "intersection", line);
}

pub fn emit_hashset_except_with(chunks: &mut [Chunk], current: usize, line: u32) {
    // `a.ExceptWith(b)` mutates in place; lower to the non-mutating ECMA
    // `difference` (returns a NEW set) then clear+refill `a` — the mutating
    // host `exceptWith` returns void, which the clear+refill pattern misreads
    // as an empty result.
    emit_hashset_mutation(chunks, current, "difference", line);
}

pub fn emit_hashset_symmetric_except_with(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_hashset_mutation(chunks, current, "symmetricDifference", line);
}

/// Set predicate (`IsSubsetOf`, `IsSupersetOf`, `Overlaps`) over any
/// `IEnumerable<T>` argument. Normalizes the argument to an ECMA Set (so
/// arrays / lists / generators work) then calls the host predicate.
/// Stack: [recv, arg] → [bool].
fn emit_hashset_predicate(chunks: &mut [Chunk], current: usize, func: &str, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    normalize_arg_to_set(chunks, current, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    call_import(chunks, current, "ecma:set", func, 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

pub fn emit_hashset_is_subset_of(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_hashset_predicate(chunks, current, "isSubsetOf", line);
}

pub fn emit_hashset_is_superset_of(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_hashset_predicate(chunks, current, "isSupersetOf", line);
}

pub fn emit_hashset_overlaps(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_hashset_predicate(chunks, current, "overlaps", line);
}

/// Composed set relations not exposed as single `ecma:set` calls:
/// `SetEquals` = `⊆ ∧ ⊇`, `IsProperSubsetOf` = `⊆ ∧ ¬⊇`,
/// `IsProperSupersetOf` = `⊇ ∧ ¬⊆`. Stack: [recv, arg] → [bool].
fn emit_hashset_relation(chunks: &mut [Chunk], current: usize, line: u32, rel: &str) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    normalize_arg_to_set(chunks, current, src, line);
    let sub_slot = chunks[current].alloc_scratch(2);
    let sup_slot = sub_slot + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    call_import(chunks, current, "ecma:set", "isSubsetOf", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sub_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    call_import(chunks, current, "ecma:set", "isSupersetOf", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sup_slot, line);

    // (a && (negate ? !b : b)), lowered as `a ? <b-term> : false`.
    let (a_slot, b_slot, negate_b) = match rel {
        "properSubset" => (sub_slot, sup_slot, true),
        "properSuperset" => (sup_slot, sub_slot, true),
        _ => (sub_slot, sup_slot, false), // setEquals
    };
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    if negate_b {
        vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    }
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_end(line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

pub fn emit_hashset_set_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_hashset_relation(chunks, current, line, "setEquals");
}

pub fn emit_hashset_is_proper_subset_of(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_hashset_relation(chunks, current, line, "properSubset");
}

pub fn emit_hashset_is_proper_superset_of(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_hashset_relation(chunks, current, line, "properSuperset");
}

pub fn emit_sorted_dictionary_entries(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let arr = base;
    let result = chunks[current].alloc_scratch(5);
    let i = result + 1;
    let j = result + 2;
    let len = result + 3;
    let key = result + 4;

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    collections::emit_clone(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);

    let inner_block = chunks[current].emit_block(line);
    let (inner_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    collections::emit_get(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    collections::emit_get(chunks, current, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(inner_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

/// `SortedSet<T>.ElementsSorted()` — spread the `ecma:set` receiver to an array
/// and sort it ascending via the shared sorted core. Backs the `foreach`
/// rewrite so iteration observes ascending order. Stack: `[set]` -> `[array]`.
pub fn emit_sorted_set_elements(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
}

/// `SortedSet<T>.Min` / `.Max` — spread the `ecma:set` receiver to a sorted array
/// via the shared sorted core and take the first / last element. (LINQ's
/// `Min`/`Max` return null over an `ecma:set`, so the ordered reads are adapted
/// explicitly.) Stack: `[set]` -> `[element]`.
pub fn emit_sorted_set_min(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_iter_values(chunks, current, line);
    vybe_emitter::sorted_collection::emit_sorted_end(chunks, current, false, line);
}

pub fn emit_sorted_set_max(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_iter_values(chunks, current, line);
    vybe_emitter::sorted_collection::emit_sorted_end(chunks, current, true, line);
}

/// `SortedSet<T>.GetViewBetween(low, high)` — spread the `ecma:set` receiver to a
/// sorted array, take the inclusive `[low, high]` range via the shared sorted
/// core, then rebuild the view as an `ecma:set` so the view's own methods
/// (`Count`/`Min`/`Max`/...) resolve through the set surface.
/// Stack: `[set, low, high]` -> `[set_view]`.
pub fn emit_sorted_set_view_between(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let recv = base;
    let low = base + 1;
    let high = base + 2;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, low, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, high, line);
    // mode 0 = both bounds; inclusive upper for .NET GetViewBetween.
    vybe_emitter::sorted_collection::emit_sorted_set_range_view(chunks, current, 0, true, line);
    call_import(chunks, current, "ecma:set", "new", 1, line);
}

pub fn emit_linked_list_add_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    collections::emit_insert_at(chunks, current, line);
}

pub fn emit_linked_list_add_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_linked_list_node_from_index(
    chunks: &mut [Chunk],
    current: usize,
    list_slot: u16,
    index_slot: u16,
    include_next: bool,
    line: u32,
) {
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let value_key = chunks[current].add_constant(Value::String(Arc::from("Value")));
    chunks[current].emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunks[current].emit_op(Op::DROP, line);

    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let lower_value_key = chunks[current].add_constant(Value::String(Arc::from("value")));
    chunks[current].emit_op_u16(Op::STRUCT_SET, lower_value_key, line);
    chunks[current].emit_op(Op::DROP, line);

    if include_next {
        let next_index_slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, next_index_slot, line);

        core_wasm::dup(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, next_index_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
        collections::emit_len(chunks, current, line);
        vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
        chunks[current].emit_if(line);
        emit_linked_list_node_from_index(chunks, current, list_slot, next_index_slot, false, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_end(line);
        let next_key = chunks[current].add_constant(Value::String(Arc::from("Next")));
        chunks[current].emit_op_u16(Op::STRUCT_SET, next_key, line);
        chunks[current].emit_op(Op::DROP, line);
    }
}

pub fn emit_linked_list_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let list = base;
    let index_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    emit_linked_list_node_from_index(chunks, current, list, index_slot, true, line);
    chunks[current].emit_end(line);
}

pub fn emit_linked_list_find(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let needle = base + 1;
    let index_slot = chunks[current].alloc_scratch(2);
    let value_slot = index_slot + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
    collections::emit_index_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let value_key = chunks[current].add_constant(Value::String(Arc::from("value")));
    chunks[current].emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

pub fn emit_vb_collection_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let object = base;
    let items = base + 1;
    let keys = base + 2;

    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, object, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items, line);

    collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);

    emit_set_field(chunks, current, object, VB_COLLECTION_ITEMS, items, line);
    emit_set_field(chunks, current, object, VB_COLLECTION_KEYS, keys, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
}

pub fn emit_vb_collection_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let value = base + 1;

    emit_get_field(chunks, current, recv, VB_COLLECTION_ITEMS, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    if argc >= 3 {
        let key = base + 2;
        emit_get_field(chunks, current, recv, VB_COLLECTION_KEYS, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
        collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }

    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_vb_collection_item(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let key = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if(line);
    emit_get_field(chunks, current, recv, VB_COLLECTION_KEYS, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    emit_get_field(chunks, current, recv, VB_COLLECTION_ITEMS, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_vb_collection_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    emit_get_field(chunks, current, base, VB_COLLECTION_ITEMS, line);
    collections::emit_len(chunks, current, line);
}

pub fn emit_vb_collection_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    emit_get_field(chunks, current, base, VB_COLLECTION_KEYS, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
}

pub fn emit_vb_collection_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let key = base + 1;
    let value_slot = chunks[current].alloc_scratch(1);

    emit_get_field(chunks, current, recv, VB_COLLECTION_KEYS, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    emit_get_field(chunks, current, recv, VB_COLLECTION_ITEMS, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_remove_value(chunks, current, line);

    emit_get_field(chunks, current, recv, VB_COLLECTION_KEYS, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}
