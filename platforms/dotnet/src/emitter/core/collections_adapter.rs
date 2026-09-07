use vybe_compiler::primitives::class_slots::{self, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::{collections, sets};

use super::object_fields::field_slot;

const VB_COLLECTION_ITEMS: &str = "__dotnet_vb_collection_items";
const VB_COLLECTION_KEYS: &str = "__dotnet_vb_collection_keys";
const BLOCKING_ITEMS: &str = "__dotnet_blocking_items";
const BLOCKING_CAPACITY: &str = "__dotnet_blocking_capacity";
const BLOCKING_COMPLETED: &str = "__dotnet_blocking_completed";
const BLOCKING_LIFO: &str = "__dotnet_blocking_lifo";
const OBSERVABLE_ITEMS: &str = "__dotnet_observable_items";
const OBSERVABLE_NOTIFYING: &str = "__dotnet_observable_notifying";
const LIST_CAPACITY: &str = "__dotnet_list_capacity";
const DICT_IGNORE_CASE: &str = "__dotnet_dict_ignorecase";

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
    chunks[current].emit_call(idx, argc, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

fn emit_sorted_comparator_get(chunks: &mut [Chunk], current: usize, source: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    chunks[current].emit_string_const(
        vybe_compiler::primitives::sorted_collection::COMPARATOR_KEY,
        line,
    );
    call_import(chunks, current, "ecma:object", "get", 2, line);
}

pub fn emit_sorted_map_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map, line);
    emit_sorted_comparator_get(chunks, current, map, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_string_const(super::comparer_adapter::COMPARER_DEFAULT, line);
    chunks[current].emit_end(line);
}

fn emit_sorted_map_is_builtin_comparer(
    chunks: &mut [Chunk],
    current: usize,
    value: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_string_const(super::comparer_adapter::ORDINAL, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_string_const(super::comparer_adapter::ORDINAL_IGNORE_CASE, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_string_const(super::comparer_adapter::COMPARER_DEFAULT, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_OR, line);
}

fn emit_sorted_map_arg_is_comparer(chunks: &mut [Chunk], current: usize, arg: u16, line: u32) {
    let method = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if(line);
    emit_sorted_map_is_builtin_comparer(chunks, current, arg, line);
    chunks[current].emit_else(line);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(arg),
        &class_slots::resolve(
            &class_slots::ClassSlot::internal("Compare"),
            &class_slots::PlainNames,
        ),
        class_slots::Dest::Local(method),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(arg),
        &class_slots::resolve(
            &class_slots::ClassSlot::internal("compare"),
            &class_slots::PlainNames,
        ),
        class_slots::Dest::Local(method),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_sorted_comparator_set(
    chunks: &mut [Chunk],
    current: usize,
    target: u16,
    comparator: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    chunks[current].emit_string_const(
        vybe_compiler::primitives::sorted_collection::COMPARATOR_KEY,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparator, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_copy_sorted_comparator(
    chunks: &mut [Chunk],
    current: usize,
    source: u16,
    target: u16,
    line: u32,
) {
    let comparator = chunks[current].alloc_scratch(1);
    emit_sorted_comparator_get(chunks, current, source, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, comparator, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparator, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_sorted_comparator_set(chunks, current, target, comparator, line);
    chunks[current].emit_end(line);
}

fn emit_string_const(chunks: &mut [Chunk], current: usize, value: &str, line: u32) {
    chunks[current].emit_string_const(value, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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

/// `SortedSet<T>()` / `SortedSet<T>(IEnumerable<T>)` / `SortedSet<T>(IComparer<T>)`.
///
/// ⛔ONE registration, branching on `argc`. `ClassType::with_constructor`
/// OVERWRITES — a second call replaces the first rather than adding an overload
/// — so declaring arity 0 and arity 1 leaves only the arity-1 body, and it then
/// runs for the no-argument call with nothing of its own on the stack.
///
/// ⛔The single argument is an ENUMERABLE or a COMPARER, .NET's two arity-1
/// overloads, and they are disjoint at run time: a comparer is never an array.
/// Dropping it unconditionally made `SortedSet[int]::new(@(50,10,90,30))` the
/// EMPTY set, so `Min`, `Max`, `Remove`, `UnionWith` and every set-algebra
/// answer came off nothing — most of the dir failed on one construction, not on
/// its own member.
pub fn emit_set_new_ignore_comparer(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        sets::emit_new(chunks, current, line);
        return;
    }
    let arg = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arg, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    collections::emit_spread_iterable_for_constructor(chunks, current, line);
    sets::emit_from_iterable(chunks, current, line);
    chunks[current].emit_else(line);
    // A comparer: the set starts empty and the ordering is the default one,
    // which is what every ordered read here already applies.
    sets::emit_new(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `List<T>()` / `ObservableCollection<T>(iterable)` — see the `argc` note on
/// [`emit_set_new_ignore_comparer`]. With no argument there is nothing to
/// spread and the collection starts empty.
pub fn emit_list_new_from_iterable(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_array_new_fixed(0, 0, line);
        return;
    }
    // ⛔`Queue<T>(int capacity)` and `Queue<T>(IEnumerable<T>)` are both arity 1
    // — the same split `HashSet<T>` has, and `Stack`, `List` and `ArrayList`
    // carry it too. A number is never an `IEnumerable` in .NET, and a capacity
    // contributes no elements, so `isFinite` separates them exactly as the
    // compiler's overload resolution does. Spreading it instead threw.
    let arg = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arg, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    let is_number = chunks[current].add_import("ecma:number", "isFinite");
    chunks[current].emit_call(is_number, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    collections::emit_spread_iterable_for_constructor(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_list_set_capacity_to_len(chunks: &mut [Chunk], current: usize, list_slot: u16, line: u32) {
    let len_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    emit_set_field(chunks, current, list_slot, LIST_CAPACITY, len_slot, line);
}

pub fn emit_list_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        let list_slot = chunks[current].alloc_scratch(1);
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, list_slot, line);
        let cap_slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, cap_slot, line);
        emit_set_field(chunks, current, list_slot, LIST_CAPACITY, cap_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
        return;
    }

    let arg_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arg_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg_slot, line);
    call_import(chunks, current, "wasm:js-number", "test", 1, line);
    chunks[current].emit_if(line);

    let numeric_list_slot = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, numeric_list_slot, line);
    emit_set_field(
        chunks,
        current,
        numeric_list_slot,
        LIST_CAPACITY,
        arg_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, numeric_list_slot, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arg_slot, line);
    // One value pushed just above, so this is the one-argument shape.
    emit_list_new_from_iterable(chunks, current, 1, line);
    let iterable_list_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, iterable_list_slot, line);
    emit_list_set_capacity_to_len(chunks, current, iterable_list_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, iterable_list_slot, line);

    chunks[current].emit_end(line);
}

pub fn emit_list_capacity(chunks: &mut [Chunk], current: usize, line: u32) {
    let list_slot = chunks[current].alloc_scratch(1);
    let cap_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, list_slot, line);
    emit_get_field(chunks, current, list_slot, LIST_CAPACITY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cap_slot, line);
    emit_slot_is_nullish(chunks, current, cap_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cap_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_list_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    let list_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let cap_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, list_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    emit_get_field(chunks, current, list_slot, LIST_CAPACITY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cap_slot, line);
    emit_slot_is_nullish(chunks, current, cap_slot, line);
    chunks[current].emit_if(line);
    emit_set_field(chunks, current, list_slot, LIST_CAPACITY, len_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cap_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_set_field(chunks, current, list_slot, LIST_CAPACITY, len_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_list_ensure_capacity(chunks: &mut [Chunk], current: usize, line: u32) {
    let desired_slot = chunks[current].alloc_scratch(1);
    let list_slot = chunks[current].alloc_scratch(1);
    let cap_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, desired_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, list_slot, line);

    emit_get_field(chunks, current, list_slot, LIST_CAPACITY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cap_slot, line);
    emit_slot_is_nullish(chunks, current, cap_slot, line);
    chunks[current].emit_if(line);
    emit_set_field(
        chunks,
        current,
        list_slot,
        LIST_CAPACITY,
        desired_slot,
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, desired_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cap_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_set_field(
        chunks,
        current,
        list_slot,
        LIST_CAPACITY,
        desired_slot,
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    emit_list_capacity(chunks, current, line);
}

pub fn emit_list_trim_excess(chunks: &mut [Chunk], current: usize, line: u32) {
    let list_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, list_slot, line);
    emit_list_set_capacity_to_len(chunks, current, list_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_readonly_observable_collection_new(_chunks: &mut [Chunk], _current: usize, _line: u32) {
    // ReadOnlyObservableCollection is a view over the source ObservableCollection:
    // preserve the backing object so Count/Item and CollectionChanged stay in sync.
}

pub fn emit_property_changed_event_args_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let property_name = stash_args(chunks, current, 1, line);
    let args = chunks[current].alloc_scratch(1);
    class_slots::emit_class_alloc(&mut chunks[current], line);
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
    class_slots::emit_class_alloc(&mut chunks[current], line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, null_value, line);
    emit_set_field(chunks, current, args, "OldItems", null_value, line);
    emit_set_field(chunks, current, args, "olditems", null_value, line);
    emit_set_field(chunks, current, args, "NewItems", null_value, line);
    emit_set_field(chunks, current, args, "newitems", null_value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, args, line);
}

pub fn emit_set_new_from_iterable(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // ⛔§24.2.1.1 step 3: a Set constructor whose iterable is `undefined` or
    // `null` yields an EMPTY set and never reaches `GetIterator`, which is what
    // `emit_spread_iterable_for_constructor` implements.
    //
    // ⛔But NO ARGUMENT is not `undefined` — nothing is pushed at all, so the
    // guard would read whatever the stack already held. `ClassType::with_constructor`
    // OVERWRITES, so the arity-0 `sets.new` registration beside this one never
    // existed and every `new HashSet<T>()` arrived here; the drain then tried to
    // iterate a stray value and threw a TypeError once it became spec-correct
    // about non-iterables. One registration, branching on `argc`.
    if argc == 0 {
        sets::emit_new(chunks, current, line);
        return;
    }
    // ⛔`HashSet<T>(int capacity)` and `HashSet<T>(IEnumerable<T>)` are both
    // arity 1. A number is never an `IEnumerable` in .NET, so the runtime test
    // separates them exactly as the compiler's overload resolution does — and
    // a capacity contributes no elements, so it yields the empty set.
    let arg = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arg, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    // `Number.isFinite` does NOT coerce — it is true for a number and false for
    // an array, a string or a generator, which is the discrimination wanted.
    let is_number = chunks[current].add_import("ecma:number", "isFinite");
    chunks[current].emit_call(is_number, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    sets::emit_new(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    collections::emit_spread_iterable_for_constructor(chunks, current, line);
    sets::emit_from_iterable(chunks, current, line);
    chunks[current].emit_end(line);
}

/// Rewrite the ECMA set in `slot` so its INSERTION ORDER is sorted order.
///
/// A `SortedSet<T>` enumerates in order, and every consumer of the backing —
/// `foreach`, a pipeline, `ToArray`, interpolation — reads the set's own
/// iteration order. Keeping the backing sorted is what makes all of them agree
/// without each one re-sorting.
fn emit_resort_set(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    let arr = chunks[current].alloc_scratch(2);
    let cursor = arr + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);
    emit_copy_sorted_comparator(chunks, current, slot, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    vybe_compiler::primitives::sorted_collection::emit_sort_if_ordered(chunks, current, arr, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    sets::emit_clear(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    sets::emit_add(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `SortedSet<T>.Add(value)` — the set add, then the backing back into order.
/// Stack: `[set, value]` → `[bool]`.
pub fn emit_sorted_set_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = chunks[current].alloc_scratch(2);
    let changed = recv + 1;
    let value = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    sets::emit_add_changed(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, changed, line);
    emit_resort_set(chunks, current, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, changed, line);
}

/// `SortedSet<T>(…)` — the `HashSet` constructor split, with the backing left
/// in sorted order. See [`emit_set_new_ignore_comparer`] for the comparer arm.
pub fn emit_sorted_set_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        sets::emit_new(chunks, current, line);
    } else {
        let arg = chunks[current].alloc_scratch(2);
        let recv = arg + 1;
        chunks[current].emit_op_u16(Op::LOCAL_SET, arg, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
        collections::emit_spread_iterable_for_constructor(chunks, current, line);
        sets::emit_from_iterable(chunks, current, line);
        chunks[current].emit_else(line);
        sets::emit_new(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
        call_import(chunks, current, "ecma:value", "typeof", 1, line);
        chunks[current].emit_string_const("function", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        emit_sorted_comparator_set(chunks, current, recv, arg, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_end(line);
    }
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    emit_resort_set(chunks, current, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
}

/// `HashSet<T>.Add` — a VALUE-equality set: a tuple, a string or a number is
/// its contents, a class instance is itself. The snapshot key renders exactly
/// that, so `hs.Add((1, 2))` twice adds one element while two distinct objects
/// stay distinct.
pub fn emit_hashset_add(chunks: &mut [Chunk], current: usize, line: u32) {
    sets::emit_add_snapshot(chunks, current, line);
}

pub fn emit_hashset_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    sets::emit_has_snapshot(chunks, current, line);
}

pub fn emit_hashset_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    sets::emit_delete_snapshot(chunks, current, line);
}

/// Rewrite the map in `map_slot` so its INSERTION ORDER is comparator order.
///
/// A `SortedDictionary<K,V>` enumerates by key, and every untyped read of the
/// backing — `$sd.Keys`, `$sd.Values`, a `foreach`, a pipeline — sees the map's
/// own iteration order rather than the sorted views. Keeping the backing
/// ordered is what makes all of them agree.
fn emit_sorted_map_reorder(chunks: &mut [Chunk], current: usize, map_slot: u16, line: u32) {
    let keys = chunks[current].alloc_scratch(3);
    let values = keys + 1;
    let cursor = keys + 2;

    emit_sorted_map_materialized_keys(chunks, current, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, values, line);
    emit_index_loop(chunks, current, cursor, keys, line, |chunks, current| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, values, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        call_import(chunks, current, "ecma:map", "get", 2, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    call_import(chunks, current, "ecma:map", "clear", 1, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_index_loop(chunks, current, cursor, keys, line, |chunks, current| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, values, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        call_import(chunks, current, "ecma:map", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
    });
}

fn emit_sorted_map_materialized_keys(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    line: u32,
) {
    let keys = chunks[current].alloc_scratch(1);
    let comparator = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    call_import(chunks, current, "ecma:map", "keys", 1, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);

    emit_sorted_comparator_get(chunks, current, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, comparator, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparator, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    emit_sorted_map_sort_keys_by_comparer(chunks, current, keys, comparator, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_string_const(
        vybe_compiler::primitives::sorted_collection::DESCENDING_MAP_KEY,
        line,
    );
    call_import(chunks, current, "ecma:object", "get", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    collections::emit_reverse(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
}

fn emit_sorted_map_sort_keys_by_comparer(
    chunks: &mut [Chunk],
    current: usize,
    keys: u16,
    comparator: u16,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(7);
    let (len, i, j, key, lhs, should_swap, lhs_index) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
    );

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);

    let inner_block = chunks[current].emit_block(line);
    let (inner_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lhs, line);

    emit_sorted_map_compare_slots(chunks, current, comparator, lhs, key, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, should_swap, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, should_swap, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lhs_index, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lhs_index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lhs, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(inner_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);
}

/// `for (cursor = 0; cursor < array.length; cursor++) { body }`.
fn emit_index_loop(
    chunks: &mut [Chunk],
    current: usize,
    cursor: u16,
    array: u16,
    line: u32,
    body: impl FnOnce(&mut [Chunk], usize),
) {
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array, line);
    collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    body(chunks, current);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `SortedDictionary<K,V>.Add(key, value)` — the map insert. Ordered APIs sort
/// the view so tuple keys remain their raw value in the backing map.
///
/// ⛔A DUPLICATE KEY THROWS. `Add` is not the indexer: .NET raises
/// `ArgumentException` rather than overwriting. Stack: `[map, key, value]`.
pub fn emit_sorted_map_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let (map, key, value) = (base, base + 1, base + 2);
    let found = chunks[current].alloc_scratch(1);

    emit_dotnet_non_null_key(chunks, current, key, line);
    emit_sorted_map_find_key(chunks, current, map, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    crate::emitter::core::exceptions::emit_new_typed(
        chunks,
        current,
        "ArgumentException",
        class_slots::ValueSource::ConstStr(
            "An item with the same key has already been added.".to_string(),
        ),
        line,
    );
    vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
        &mut chunks[current],
        "ArgumentException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

/// `SortedDictionary<K,V>.Item(key) = value` — overwrite/insert. Ordered APIs
/// sort the view so tuple keys remain their raw value in the backing map.
pub fn emit_sorted_map_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let (map, key, value) = (base, base + 1, base + 2);
    let actual_key = chunks[current].alloc_scratch(1);
    emit_dotnet_non_null_key(chunks, current, key, line);
    emit_sorted_map_find_key(chunks, current, map, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, actual_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, actual_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, actual_key, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
}

pub fn emit_sorted_map_contains_key(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let (map, key) = (base, base + 1);
    emit_dotnet_non_null_key(chunks, current, key, line);
    emit_sorted_map_find_key(chunks, current, map, key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_sorted_map_get_or_throw(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let (map, key) = (base, base + 1);
    let actual_key = chunks[current].alloc_scratch(1);
    emit_dotnet_non_null_key(chunks, current, key, line);
    emit_sorted_map_find_key(chunks, current, map, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, actual_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, actual_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "KeyNotFoundException",
        "The given key was not present in the dictionary.",
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, actual_key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_end(line);
}

pub fn emit_sorted_map_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let (map, key) = (base, base + 1);
    let actual_key = chunks[current].alloc_scratch(1);
    emit_dotnet_non_null_key(chunks, current, key, line);
    emit_sorted_map_find_key(chunks, current, map, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, actual_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, actual_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, actual_key, line);
    call_import(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_end(line);
}

/// `SortedDictionary<K,V>(…)` — the shared sorted-collection construction, plus
/// the `IDictionary` overload.
///
/// ⛔ONLY A CALLABLE ARGUMENT BECOMES THE COMPARATOR. The sorted core's key
/// view guards its stored comparator with a NULL test, not a callable test, so
/// an `IDictionary` parked under that key is then invoked as the sort predicate
/// and every ordered read traps. A non-callable argument is drained for entries
/// instead — `ecma:map.keys` answers empty for anything that is not a map, so
/// the two shapes need no test between them beyond callability.
pub fn emit_sorted_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        vybe_compiler::primitives::sorted_collection::emit_sorted_collection_new(
            chunks, current, 0, true, line,
        );
        return;
    }
    let arg = chunks[current].alloc_scratch(4);
    let map = arg + 1;
    let keys = arg + 2;
    let cursor = arg + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, arg, line);
    emit_sorted_map_arg_is_comparer(chunks, current, arg, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    vybe_compiler::primitives::sorted_collection::emit_sorted_collection_new(
        chunks, current, 1, true, line,
    );
    chunks[current].emit_else(line);
    vybe_compiler::primitives::sorted_collection::emit_sorted_collection_new(
        chunks, current, 0, true, line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    call_import(chunks, current, "ecma:map", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    emit_index_loop(chunks, current, cursor, keys, line, |chunks, current| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        call_import(chunks, current, "ecma:map", "get", 2, line);
        call_import(chunks, current, "ecma:map", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    emit_sorted_map_reorder(chunks, current, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
}

pub fn emit_dict_new_ignore_arg(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let map = chunks[current].alloc_scratch(1);
    call_import(chunks, current, "ecma:map", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map, line);

    for offset in 0..argc as u16 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + offset, line);
        chunks[current].emit_string_const(super::comparer_adapter::ORDINAL_IGNORE_CASE, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
        chunks[current].emit_string_const(DICT_IGNORE_CASE, line);
        chunks[current].emit_bool_const(true, line);
        call_import(chunks, current, "ecma:object", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
}

fn emit_dotnet_dict_key(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    key_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "ArgumentNullException",
        "Value cannot be null. (Parameter 'key')",
        line,
    );
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_string_const(DICT_IGNORE_CASE, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    chunks[current].emit_bool_const(true, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    vybe_compiler::primitives::dict::emit_canonical_key_slot(
        chunks, current, map_slot, key_slot, line,
    );
}

fn emit_dotnet_non_null_key(chunks: &mut [Chunk], current: usize, key_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "ArgumentNullException",
        "Value cannot be null. (Parameter 'key')",
        line,
    );
    chunks[current].emit_end(line);
}

fn emit_sorted_map_compare_slots(
    chunks: &mut [Chunk],
    current: usize,
    comparator: u16,
    left: u16,
    right: u16,
    line: u32,
) {
    let abi = vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    let method = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, comparator, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_sorted_map_default_compare_slots(chunks, current, left, right, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, comparator, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparator, line);
    let recv = vybe_compiler::primitives::callable::emit_callback_receiver(
        &mut chunks[current],
        abi,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        2 + recv,
        line,
    );
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, comparator, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparator, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    super::comparer_adapter::emit_compare(chunks, current, line);
    chunks[current].emit_else(line);

    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(comparator),
        &class_slots::resolve(
            &class_slots::ClassSlot::internal("Compare"),
            &class_slots::PlainNames,
        ),
        class_slots::Dest::Local(method),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    let recv = vybe_compiler::primitives::callable::emit_callback_receiver(
        &mut chunks[current],
        abi,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        2 + recv,
        line,
    );
    chunks[current].emit_else(line);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(comparator),
        &class_slots::resolve(
            &class_slots::ClassSlot::internal("compare"),
            &class_slots::PlainNames,
        ),
        class_slots::Dest::Local(method),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
    let recv = vybe_compiler::primitives::callable::emit_callback_receiver(
        &mut chunks[current],
        abi,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        2 + recv,
        line,
    );
    chunks[current].emit_else(line);
    emit_sorted_map_default_compare_slots(chunks, current, left, right, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_sorted_map_default_compare_slots(
    chunks: &mut [Chunk],
    current: usize,
    left: u16,
    right: u16,
    line: u32,
) {
    let method = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);

    let done = chunks[current].emit_block_typed(line, 1);
    let compare_slot = vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Compare);
    for name in ["compare", "CompareTo", "compareTo", compare_slot.as_str()] {
        class_slots::emit_class_get(
            &mut chunks[current],
            ObjSource::Local(left),
            &class_slots::resolve(
                &class_slots::ClassSlot::internal(name),
                &class_slots::PlainNames,
            ),
            class_slots::Dest::Local(method),
            line,
        );
        chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
        call_import(chunks, current, "ecma:value", "typeof", 1, line);
        chunks[current].emit_string_const("function", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, method, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
        vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
            &mut chunks[current],
            2,
            line,
        );
        chunks[current].emit_br(1, line);
        chunks[current].emit_end(line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(done);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_sorted_map_find_key(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    key_slot: u16,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(6);
    let (comparator, keys, cursor, found, candidate, cmp) =
        (base, base + 1, base + 2, base + 3, base + 4, base + 5);

    emit_sorted_comparator_get(chunks, current, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, comparator, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    call_import(chunks, current, "ecma:map", "keys", 1, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);

    let outer_block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, candidate, line);
    emit_sorted_map_compare_slots(chunks, current, comparator, candidate, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cmp, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cmp, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, candidate, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
}

pub fn emit_dict_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let recv = base;
    let key = base + 1;
    let value = base + 2;

    emit_dotnet_dict_key(chunks, current, recv, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    crate::emitter::core::exceptions::emit_new_typed(
        chunks,
        current,
        "ArgumentException",
        class_slots::ValueSource::ConstStr(
            "An item with the same key has already been added.".to_string(),
        ),
        line,
    );
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
        chunk,
        "ArgumentException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

pub fn emit_dict_contains_key(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let key = base + 1;

    emit_dotnet_dict_key(chunks, current, recv, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
}

pub fn emit_dict_contains_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let needle = base + 1;
    let values = chunks[current].alloc_scratch(4);
    let cursor = values + 1;
    let found = values + 2;
    let elem = values + 3;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:map", "values", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, values, line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);

    emit_index_loop(chunks, current, cursor, values, line, |chunks, current| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, values, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, elem, line);
        super::linq_adapter::emit_linq_value_equals_slots(chunks, current, elem, needle, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        core_wasm::bool_const(&mut chunks[current], line, true);
        chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);
        chunks[current].emit_end(line);
    });

    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
}

pub fn emit_dict_to_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(5);
    let entries = map + 1;
    let cursor = map + 2;
    let result = map + 3;
    let entry = map + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    call_import(chunks, current, "ecma:map", "entries", 1, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);

    emit_index_loop(chunks, current, cursor, entries, line, |chunks, current| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, entry, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_i32_const(0, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_i32_const(1, line);
        collections::emit_get(chunks, current, line);
        super::key_value_pair_adapter::emit_key_value_pair_new(chunks, current, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

pub fn emit_dict_get_value_or_default(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let key = base + 1;
    let default = base + 2;

    emit_dotnet_dict_key(chunks, current, recv, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, default, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    chunks[current].emit_end(line);
}

pub fn emit_dict_set_item(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let recv = base;
    let key = base + 1;
    let value = base + 2;

    emit_dotnet_dict_key(chunks, current, recv, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
}

pub fn emit_dict_get_or_throw(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let map = base;
    let key = base + 1;

    emit_dotnet_dict_key(chunks, current, map, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    crate::emitter::core::exceptions::emit_new_typed(
        chunks,
        current,
        "KeyNotFoundException",
        class_slots::ValueSource::ConstStr(
            "The given key was not present in the dictionary.".to_string(),
        ),
        line,
    );
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
        chunk,
        "KeyNotFoundException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
}

pub fn emit_dict_try_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let recv = base;
    let key = base + 1;
    let value = base + 2;

    emit_dotnet_dict_key(chunks, current, recv, key, line);
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

    // `ICollection<KeyValuePair<K,V>>.Remove(pair)` removes only when the key
    // AND the value match, unlike `Remove(key)`. The pair reaches this in TWO
    // shapes: a real `KeyValuePair` object from
    // `key_value_pair_adapter::emit_key_value_pair_new`, and the positional
    // `[key, value]` that `Dictionary.Entries` still yields (`ecma:map.entries`,
    // ECMA-262 §24.1.3.4). Both must be recognised.
    //
    // The object form is identified by the SHARED `__value_eq` stamp rather
    // than by probing for a `Key` property: an absent property reads back as
    // `Undefined`, and `REF_IS_NULL` is FALSE for `Undefined`, so a
    // presence test would call every argument a pair.
    let is_pair = chunks[current].alloc_scratch(1);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, is_pair, line);

    vybe_compiler::primitives::records::emit_is_value_eq(&mut chunks[current], arg, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    chunks[current].emit_string_const("Key", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    chunks[current].emit_string_const("Value", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, is_pair, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    vybe_compiler::primitives::tuples::emit_is_tuple(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, is_pair, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, is_pair, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    emit_dotnet_dict_key(chunks, current, recv, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, current_value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, current_value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    emit_dotnet_dict_key(chunks, current, recv, arg, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `System.Index` — a VALUE and a direction, which is the whole struct.
///
/// `Index(value, isFromEnd)`; `FromStart(n)` and `FromEnd(n)` are the same
/// struct with the flag supplied, so they push it rather than taking it.
pub fn emit_index_new(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::dispatch::emit_value_type_new(
        &mut chunks[current],
        "Index",
        &["Value", "IsFromEnd"],
        line,
    );
}

pub fn emit_index_from(chunks: &mut [Chunk], current: usize, from_end: bool, line: u32) {
    chunks[current].emit_bool_const(from_end, line);
    emit_index_new(chunks, current, line);
}

fn emit_index_offset_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    index_slot: u16,
    length_slot: u16,
    line: u32,
) {
    emit_get_field(chunks, current, index_slot, "IsFromEnd", line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, length_slot, line);
    emit_get_field(chunks, current, index_slot, "Value", line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_else(line);
    emit_get_field(chunks, current, index_slot, "Value", line);
    chunks[current].emit_end(line);
}

pub fn emit_index_get_offset(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    emit_index_offset_from_slots(chunks, current, base, base + 1, line);
}

pub fn emit_index_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let (left, right) = (base, base + 1);
    emit_get_field(chunks, current, left, "Value", line);
    emit_get_field(chunks, current, right, "Value", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_get_field(chunks, current, left, "IsFromEnd", line);
    emit_get_field(chunks, current, right, "IsFromEnd", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_index_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    emit_get_field(chunks, current, base, "Value", line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    emit_get_field(chunks, current, base, "IsFromEnd", line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_ADD, line);
}

pub fn emit_index_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    emit_get_field(chunks, current, base, "IsFromEnd", line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("^", line);
    emit_get_field(chunks, current, base, "Value", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    emit_get_field(chunks, current, base, "Value", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// `Tuple.Create(…)` / `ValueTuple.Create(…)` — the items become `Item1`…`ItemN`.
///
/// The arity IS the shape, so the field list is built from `argc` rather than
/// declared once per size. `Tuple` and `ValueTuple` differ only in being a
/// reference versus a value type, which nothing here observes, so both reach
/// this and carry the type name they were called on.
pub fn emit_tuple_create(
    chunks: &mut [Chunk],
    current: usize,
    type_name: &str,
    argc: u8,
    line: u32,
) {
    let _ = type_name;
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, argc as u16, line);
}

/// Push `hasOwn(recv, "Item<n>")` as an i32.
///
/// A tuple's items are contiguous from `Item1`, so a PRESENCE test walks the
/// shape without a stored arity — and unlike a value test it still sees an
/// item holding null, which .NET allows.
fn emit_tuple_has_item(chunks: &mut [Chunk], current: usize, recv: u16, n: u8, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_i32_const(n as i32, line);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn emit_tuple_item(chunks: &mut [Chunk], current: usize, recv: u16, n: u8, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_i32_const(n as i32 - 1, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
}

/// The .NET arities: `Tuple<T1>` through `Tuple<T1…T8>`.
const TUPLE_MAX_ARITY: u8 = 8;

/// `tuple.ToString()` — `(item1, item2, …)`, .NET's own rendering.
pub fn emit_tuple_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    let out = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);
    let item_text = chunks[current].alloc_scratch(1);

    chunks[current].emit_string_const("(", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    for n in 1..=TUPLE_MAX_ARITY {
        emit_tuple_has_item(chunks, current, recv, n, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        if n > 1 {
            chunks[current].emit_string_const(", ", line);
            vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        }
        emit_tuple_item(chunks, current, recv, n, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, item, line);
        super::console_adapter::emit_dotnet_stringify(&mut chunks[current], item, item_text, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item_text, line);
        vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const(")", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
}

/// `tuple.CompareTo(other)` — item by item, left to right, stopping at the
/// first pair that differs. Negative, zero or positive, as `IComparable` asks.
pub fn emit_tuple_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let (recv, other) = (base, base + 1);
    let result = chunks[current].alloc_scratch(1);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    for n in 1..=TUPLE_MAX_ARITY {
        chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        emit_tuple_has_item(chunks, current, recv, n, line);
        chunks[current].emit_if(line);
        emit_tuple_item(chunks, current, recv, n, line);
        emit_tuple_item(chunks, current, other, n, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_i32_const(-1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
        chunks[current].emit_else(line);
        emit_tuple_item(chunks, current, recv, n, line);
        emit_tuple_item(chunks, current, other, n, line);
        vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

/// `tuple.Equals(other)` — value equality for the canonical tuple shape.
pub fn emit_tuple_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    vybe_compiler::primitives::tuples::emit_tuple_value_eq(&mut chunks[current], line);
}

/// `System.Range` — a start index and an end index.
pub fn emit_range_new(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::dispatch::emit_value_type_new(
        &mut chunks[current],
        "Range",
        &["Start", "End"],
        line,
    );
}

/// `Range.StartAt(i)` — from `i` to the end; `EndAt(i)` — from the start to
/// `i`; `All` — the whole span. Each supplies the open end as the `Index` that
/// bounds it, so all three build the same struct.
pub fn emit_range_open(chunks: &mut [Chunk], current: usize, at_start: bool, line: u32) {
    if at_start {
        // [start] → [start, Index::FromEnd(0)]
        chunks[current].emit_i32_const(0, line);
        emit_index_from(chunks, current, true, line);
    } else {
        // [end] → [Index::FromStart(0), end]
        let given = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, given, line);
        chunks[current].emit_i32_const(0, line);
        emit_index_from(chunks, current, false, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, given, line);
    }
    emit_range_new(chunks, current, line);
}

/// `Range.All` — start of sequence to end of sequence.
pub fn emit_range_all(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(0, line);
    emit_index_from(chunks, current, false, line);
    chunks[current].emit_i32_const(0, line);
    emit_index_from(chunks, current, true, line);
    emit_range_new(chunks, current, line);
}

pub fn emit_range_get_offset_and_length(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let (range, length) = (base, base + 1);
    let offset = chunks[current].alloc_scratch(2);
    let end_offset = offset + 1;
    emit_get_field(chunks, current, range, "Start", line);
    let start = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    emit_index_offset_from_slots(chunks, current, start, length, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, offset, line);
    emit_get_field(chunks, current, range, "End", line);
    let end = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end, line);
    emit_index_offset_from_slots(chunks, current, end, length, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_offset, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, offset, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_offset, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, offset, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, offset, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_offset, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, offset, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    crate::emitter::dispatch::emit_value_type_new(
        &mut chunks[current],
        "ValueTuple",
        &["Item1", "Item2", "Offset", "Length"],
        line,
    );
}

pub fn emit_range_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let (left, right) = (base, base + 1);
    emit_get_field(chunks, current, left, "Start", line);
    emit_get_field(chunks, current, right, "Start", line);
    emit_index_equals(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_get_field(chunks, current, left, "End", line);
    emit_get_field(chunks, current, right, "End", line);
    emit_index_equals(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_range_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    emit_get_field(chunks, current, base, "Start", line);
    emit_index_hash_code(chunks, current, line);
    chunks[current].emit_i32_const(31, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    emit_get_field(chunks, current, base, "End", line);
    emit_index_hash_code(chunks, current, line);
    chunks[current].emit_op(Op::I32_ADD, line);
}

pub fn emit_range_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    emit_get_field(chunks, current, base, "Start", line);
    emit_index_to_string(chunks, current, line);
    chunks[current].emit_string_const("..", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    emit_get_field(chunks, current, base, "End", line);
    emit_index_to_string(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
}

/// `IsEmpty` — the concurrent collections' own spelling for "no elements".
///
/// Stack: `[collection]` → `[bool]`.
pub fn emit_collection_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `Stack<T>.ToArray()` / `ConcurrentStack<T>.ToArray()` — TOP TO BOTTOM.
///
/// .NET documents both as returning the newest element first, which is the
/// reverse of the insertion order the backing array holds. A plain clone hands
/// back insertion order, so `$s.Push("a"); $s.Push("b"); $s.ToArray()[0]`
/// answered `a` where .NET says `b`.
///
/// Stack: `[stack]` → `[array]`.
pub fn emit_stack_to_array(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::collections::emit_clone(chunks, current, line);
    vybe_compiler::primitives::collections::emit_reverse(chunks, current, line);
}

/// `Stack<T>.CopyTo` / `ConcurrentStack<T>.CopyTo` — copy top-to-bottom.
pub fn emit_stack_copy_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let source = base;
    let dest = base + 1;
    let index = base + 2;

    emit_check_copy_to_args(chunks, current, dest, index, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    emit_stack_to_array(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dest, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_copy_to(chunks, current, line);
}

/// `ICollection<T>.CopyTo(array, index)` — .NET exception shape wrapped around
/// the shared collection copy primitive.
pub fn emit_collection_copy_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let source = base;
    let dest = base + 1;
    let index = base + 2;

    emit_check_copy_to_args(chunks, current, dest, index, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dest, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_copy_to(chunks, current, line);
}

fn emit_check_copy_to_args(chunks: &mut [Chunk], current: usize, dest: u16, index: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, dest, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "ArgumentNullException",
        "Value cannot be null. (Parameter 'array')",
        line,
    );
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "ArgumentOutOfRangeException",
        "Index was out of range. Must be non-negative.",
        line,
    );
    chunks[current].emit_end(line);
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

    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    if argc > 0 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
        call_import(chunks, current, "wasm:js-number", "test", 1, line);
        chunks[current].emit_if(line);
        // A bounded capacity is at least 1. .NET rejects 0 and negatives from
        // the constructor rather than building a collection nothing fits in.
        chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
        chunks[current].emit_i32_const(1, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        emit_throw_dotnet_exception(
            chunks,
            current,
            "ArgumentOutOfRangeException",
            "The boundedCapacity argument must be positive.",
            line,
        );
        chunks[current].emit_end(line);
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

/// Throw when the collection has been marked complete for adding.
fn emit_check_adding_open(chunks: &mut [Chunk], current: usize, recv: u16, line: u32) {
    emit_blocking_field(chunks, current, recv, BLOCKING_COMPLETED, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "InvalidOperationException",
        "The collection has been marked as complete with regards to additions.",
        line,
    );
    chunks[current].emit_end(line);
}

/// `bc.TryAdd(item)` / `bc.TryAdd(item, millisecondsTimeout)`.
///
/// `argc` counts the receiver. The trailing argument is a timeout, and nothing
/// here blocks, so a full collection answers false immediately — which is what
/// a timeout that elapses answers anyway.
pub fn emit_blocking_collection_try_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc.max(2), line);
    let recv = base;
    let value = base + 1;
    emit_check_adding_open(chunks, current, recv, line);
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
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, capacity, line);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
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

/// Raise `OperationCanceledException` when the token in `slot` is cancelled.
///
/// ⛔ The check is spelled out rather than delegated to
/// `thread_adapter::emit_cancellation_token_throw_if_requested`: that helper's
/// throwing arm and its `ref.null` arm leave the operand stack at DIFFERENT
/// heights, which is harmless as a whole method body and corrupts the frame
/// when it is one step of a larger emitter — inside a `Try` the next call then
/// reads a leftover operand as its callee (`f64 is not callable`). A void `if`
/// whose only arm throws is balanced.
fn emit_check_not_cancelled(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_get_field(chunks, current, slot, "iscancellationrequested", line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    crate::emitter::core::exceptions::emit_new_typed(
        chunks,
        current,
        "OperationCanceledException",
        class_slots::ValueSource::ConstStr("The operation was canceled.".to_string()),
        line,
    );
    vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
        &mut chunks[current],
        "OperationCanceledException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// `bc.Add(item)` / `bc.Add(item, cancellationToken)`.
pub fn emit_blocking_collection_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 2 {
        let base = stash_args(chunks, current, argc, line);
        // The token is checked before the item is ever considered.
        emit_check_not_cancelled(chunks, current, base + 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    }
    emit_blocking_collection_try_add(chunks, current, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `bc.Take()` / `bc.Take(cancellationToken)` / the value-returning `TryTake`.
pub fn emit_blocking_collection_take(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc.max(1), line);
    if argc > 1 {
        emit_check_not_cancelled(chunks, current, base + 1, line);
    }
    let recv = base;
    let items = chunks[current].alloc_scratch(2);
    let lifo = items + 1;

    // An empty collection that will never be added to again cannot ever
    // answer, so .NET raises rather than blocking forever.
    emit_blocking_field(chunks, current, recv, BLOCKING_ITEMS, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_blocking_field(chunks, current, recv, BLOCKING_COMPLETED, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "InvalidOperationException",
        "The collection argument is empty and has been marked as complete with regards to additions.",
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_blocking_collection_is_completed(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    emit_blocking_field(chunks, current, recv, BLOCKING_COMPLETED, line);
}

pub fn emit_blocking_collection_items(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    emit_blocking_field(chunks, current, recv, BLOCKING_ITEMS, line);
}

/// `bc.ToArray()` — a COPY. The live array would let a later `Add` show up in
/// a snapshot taken before it.
pub fn emit_blocking_collection_to_array(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_blocking_collection_items(chunks, current, line);
    collections::emit_clone(chunks, current, line);
}

/// `bc.Dispose()` — the collection holds no host resource, so releasing it is
/// the receiver going out of scope.
pub fn emit_blocking_collection_dispose(chunks: &mut [Chunk], current: usize, line: u32) {
    stash_args(chunks, current, 1, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    class_slots::emit_class_alloc(&mut chunks[current], line);
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
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, old_items_value, line);
    emit_set_field(chunks, current, args, "OldItems", old_items_value, line);
    emit_set_field(chunks, current, args, "olditems", old_items_value, line);

    let new_items_value = chunks[current].alloc_scratch(1);
    if let Some(slot) = new_items {
        emit_array_with_slots(chunks, current, &[slot], line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    class_slots::emit_class_alloc(&mut chunks[current], line);
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
    emit_slot_is_nullish(chunks, current, delegate, line);
    chunks[current].emit_if(line);
    emit_get_field(chunks, current, recv, fallback_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, delegate, line);
    chunks[current].emit_end(line);

    emit_slot_is_nullish(chunks, current, delegate, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, args, line);
    vybe_compiler::primitives::delegates::emit_invoke(chunks, current, 3, line);
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
    crate::emitter::core::exceptions::emit_new_typed(
        chunks,
        current,
        exception_name,
        class_slots::ValueSource::ConstStr(message.to_string()),
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);

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
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `List.Clear()` — empty the ITEMS, whichever shape the receiver has.
///
/// ⛔ `collections.clear` splices the RECEIVER, which is right only when the
/// list value is a bare array. A `List` can also be an object carrying its
/// elements in `__dotnet_observable_items` (see `emit_observable_items_slot`),
/// and splicing that object left the elements untouched — `Count` still
/// answered 2 after `Clear()`, and the receiver came back not-a-number on the
/// next read. This routes through the same items slot every other List member
/// uses, so both shapes clear.
///
/// Distinct from `emit_observable_collection_clear`, which additionally raises
/// the `Reset` change notification an `ObservableCollection` owes its
/// subscribers; a plain `List` has none.
pub fn emit_list_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash_args(chunks, current, 1, line);
    let items = emit_observable_items_slot(chunks, current, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    collections::emit_clear(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_observable_collection_on_changed(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let args = base + 1;
    emit_observable_collection_changed(chunks, current, recv, args, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// The `IEnumerable<T>` argument of a set operation, as an ECMA Set. Stores the
/// Set back in `slot`.
///
/// ⛔A TAG COMPARISON IS NOT A TYPE TEST HERE. `ecma:object.toStringTag` reads
/// the `__type` property ahead of the VM kind, and every set this platform
/// builds carries `__type = "HashSet"` / `"SortedSet"`, so a genuine Set answers
/// `[object HashSet]`, never `[object Set]`. Any pass-through guard written on
/// that comparison misses, the argument falls to a drain that yields nothing,
/// and set algebra with a SET argument becomes a silent no-op.
///
/// So there is no guard: `iterForOf` dispatches on the VM's own `ObjectKind`,
/// which `__type` cannot shadow, and drains Array, Map, Set and TypedArray
/// alike. A fresh Set is built from the elements — a .NET set operation never
/// mutates its argument, so the copy costs only the walk.
fn normalize_arg_to_set(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_iter_for_of(chunks, current, line);
    sets::emit_from_iterable_snapshot(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn emit_hashset_mutation(chunks: &mut [Chunk], current: usize, func: &str, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    let values = chunks[current].alloc_scratch(3);
    let cursor = values + 1;
    let elem = values + 2;

    // C# set methods take any `IEnumerable<T>`; the ECMA set operations require
    // a set-like operand. Normalize the argument to an ECMA Set via
    // `new Set(iterable)` (§24.2.1.1) so arrays / lists / generators all work —
    // pure adaptation, no host or VM change.
    normalize_arg_to_set(chunks, current, src, line);

    if func == "intersection" {
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    }
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, values, line);

    emit_index_loop(chunks, current, cursor, values, line, |chunks, current| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, values, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, elem, line);

        match func {
            "intersection" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
                sets::emit_has_snapshot(chunks, current, line);
                vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
                chunks[current].emit_if_value(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
                sets::emit_delete_snapshot(chunks, current, line);
                chunks[current].emit_op(Op::DROP, line);
                chunks[current].emit_end(line);
            }
            "difference" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
                sets::emit_delete_snapshot(chunks, current, line);
                chunks[current].emit_op(Op::DROP, line);
            }
            "symmetricDifference" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
                sets::emit_has_snapshot(chunks, current, line);
                vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
                chunks[current].emit_if_value(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
                sets::emit_delete_snapshot(chunks, current, line);
                chunks[current].emit_op(Op::DROP, line);
                chunks[current].emit_else(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
                sets::emit_add_snapshot(chunks, current, line);
                chunks[current].emit_op(Op::DROP, line);
                chunks[current].emit_end(line);
            }
            _ => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
                sets::emit_add_snapshot(chunks, current, line);
                chunks[current].emit_op(Op::DROP, line);
            }
        }
    });

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_hashset_union_with(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_hashset_mutation(chunks, current, "union", line);
}

pub fn emit_hashset_intersect_with(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_hashset_mutation(chunks, current, "intersection", line);
}

pub fn emit_hashset_except_with(chunks: &mut [Chunk], current: usize, line: u32) {
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
    match func {
        "isSubsetOf" => sets::emit_is_subset_of(chunks, current, line),
        "isSupersetOf" => sets::emit_is_superset_of(chunks, current, line),
        _ => {
            sets::emit_is_disjoint_from(chunks, current, line);
            vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
        }
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    sets::emit_is_subset_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sub_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    sets::emit_is_superset_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sup_slot, line);

    // (a && (negate ? !b : b)), lowered as `a ? <b-term> : false`.
    let (a_slot, b_slot, negate_b) = match rel {
        "properSubset" => (sub_slot, sup_slot, true),
        "properSuperset" => (sup_slot, sub_slot, true),
        _ => (sub_slot, sup_slot, false), // setEquals
    };
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    if negate_b {
        vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    }
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    emit_sorted_map_entries(chunks, current, line);
}

pub fn emit_sorted_map_keys(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let cursor = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    emit_sorted_map_entries(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    emit_index_loop(chunks, current, cursor, entries, line, |chunks, current| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_i32_const(0, line);
        collections::emit_get(chunks, current, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_sorted_map_values(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let cursor = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    emit_sorted_map_entries(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    emit_index_loop(chunks, current, cursor, entries, line, |chunks, current| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_i32_const(1, line);
        collections::emit_get(chunks, current, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_sorted_list_remove_at(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let map = base;
    let index = base + 1;
    let keys = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    emit_sorted_map_keys(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_get(chunks, current, line);
    call_import(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_sorted_list_index_of_key(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let map = base;
    let key = base + 1;
    let keys = chunks[current].alloc_scratch(1);
    let actual_key = chunks[current].alloc_scratch(1);

    emit_sorted_map_find_key(chunks, current, map, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, actual_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, actual_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    emit_sorted_map_keys(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, actual_key, line);
    collections::emit_index_of(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_sorted_list_index_of_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let map = base;
    let value = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    emit_sorted_map_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_index_of(chunks, current, line);
}

pub fn emit_sorted_map_entries(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    call_import(chunks, current, "ecma:map", "entries", 1, line);
    collections::emit_iter_values(chunks, current, line);
    let entries = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries, line);
    emit_sorted_map_sort_entries_by_key(chunks, current, map, entries, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
}

pub fn emit_sorted_map_copy_to(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (map, dest, index) = if argc >= 2 {
        let base = stash_args(chunks, current, 3, line);
        (base + 2, base, base + 1)
    } else {
        let base = stash_args(chunks, current, 2, line);
        let index = chunks[current].alloc_scratch(1);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
        (base + 1, base, index)
    };
    let entries = chunks[current].alloc_scratch(4);
    let cursor = entries + 1;
    let entry = entries + 2;
    let pair = entries + 3;

    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    emit_sorted_map_entries(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);

    emit_index_loop(chunks, current, cursor, entries, line, |chunks, current| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, entry, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_i32_const(0, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_i32_const(1, line);
        collections::emit_get(chunks, current, line);
        super::key_value_pair_adapter::emit_key_value_pair_new(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, pair, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, dest, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
        collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn emit_sorted_map_entry_key(chunks: &mut [Chunk], current: usize, entry_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, entry_slot, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
}

fn emit_sorted_map_sort_entries_by_key(
    chunks: &mut [Chunk],
    current: usize,
    map: u16,
    entries: u16,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(9);
    let (comparator, len, i, j, key_entry, lhs_entry, should_swap, lhs_key, key_key) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
        base + 7,
        base + 8,
    );

    emit_sorted_comparator_get(chunks, current, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, comparator, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_entry, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);

    let inner_block = chunks[current].emit_block(line);
    let (inner_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lhs_entry, line);

    emit_sorted_map_entry_key(chunks, current, lhs_entry, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lhs_key, line);
    emit_sorted_map_entry_key(chunks, current, key_entry, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_key, line);
    emit_sorted_map_compare_slots(chunks, current, comparator, lhs_key, key_key, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, should_swap, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, should_swap, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lhs_entry, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(inner_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_entry, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_string_const(
        vybe_compiler::primitives::sorted_collection::DESCENDING_MAP_KEY,
        line,
    );
    call_import(chunks, current, "ecma:object", "get", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
    collections::emit_reverse(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

/// `SortedSet<T>.ElementsSorted()` — spread the `ecma:set` receiver to an array
/// and sort it ascending via the shared sorted core. Backs the `foreach`
/// rewrite so iteration observes ascending order. Stack: `[set]` -> `[array]`.
pub fn emit_sorted_set_elements(chunks: &mut [Chunk], current: usize, line: u32) {
    let set = chunks[current].alloc_scratch(2);
    let arr = set + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, set, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);
    emit_copy_sorted_comparator(chunks, current, set, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    vybe_compiler::primitives::sorted_collection::emit_sort_if_ordered(chunks, current, arr, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
}

/// `SortedSet<T>.Min` / `.Max` — spread the `ecma:set` receiver to a sorted array
/// via the shared sorted core and take the first / last element. (LINQ's
/// `Min`/`Max` return null over an `ecma:set`, so the ordered reads are adapted
/// explicitly.) Stack: `[set]` -> `[element]`.
pub fn emit_sorted_set_min(chunks: &mut [Chunk], current: usize, line: u32) {
    let set = chunks[current].alloc_scratch(2);
    let arr = set + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, set, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);
    emit_copy_sorted_comparator(chunks, current, set, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    vybe_compiler::primitives::sorted_collection::emit_sorted_end(chunks, current, false, line);
}

pub fn emit_sorted_set_max(chunks: &mut [Chunk], current: usize, line: u32) {
    let set = chunks[current].alloc_scratch(2);
    let arr = set + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, set, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);
    emit_copy_sorted_comparator(chunks, current, set, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    vybe_compiler::primitives::sorted_collection::emit_sorted_end(chunks, current, true, line);
}

/// `SortedSet<T>.GetViewBetween(low, high)` — spread the `ecma:set` receiver to a
/// sorted array, take the inclusive `[low, high]` range via the shared sorted
/// core, then rebuild the view as an `ecma:set` so the view's own methods
/// (`Count`/`Min`/`Max`/...) resolve through the set surface.
///
/// ⛔THE VIEW CARRIES `__type = "SortedSet"`. A language that types its locals
/// (C#) reaches the view's members through the declared type, but an untyped
/// receiver resolves on the stamp alone, and a bare set answers nothing for
/// `$view.Min`.
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
    vybe_compiler::primitives::sorted_collection::emit_sorted_set_range_view(
        chunks, current, 0, true, line,
    );
    sets::emit_from_iterable(chunks, current, line);
    core_wasm::dup(&mut chunks[current], line);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Stack,
        &field_slot("__type"),
        ValueSource::ConstStr("SortedSet".to_string()),
        line,
    );
}

/// `list.AddFirst(v)` / `list.AddLast(v)` — insert, and answer the NODE that
/// now holds the value, which is what `$n = $ll.AddLast(20)` binds.
pub fn emit_linked_list_add_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let index = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    collections::emit_insert_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    emit_node_at(chunks, current, base, index, line);
}

pub fn emit_linked_list_add_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let index = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    emit_node_at(chunks, current, base, index, line);
}

/// The list's elements as `LinkedListNode<T>` objects, doubly linked.
///
/// ⛔A NODE IS NOT A VALUE WRAPPER. The corpus walks `$node.Next.Value`,
/// `$node.Previous.Value` and `$node.List`, so a node built one hop deep ends
/// the walk at the second. The whole chain is materialized instead, in two
/// passes — every node exists before any `Next`/`Previous` can point at one.
///
/// ⛔EVERY FIELD IS WRITTEN TWICE, IN BOTH SPELLINGS. PowerShell folds a member
/// name to lower case before resolving it while other readers keep the .NET
/// spelling, and one slot cannot answer both.
///
/// Leaves the node array in the returned slot.
fn emit_node_chain(chunks: &mut [Chunk], current: usize, list: u16, line: u32) -> u16 {
    let nodes = chunks[current].alloc_scratch(4);
    let cursor = nodes + 1;
    let entry = nodes + 2;
    let scratch = nodes + 3;

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nodes, line);

    emit_index_loop(chunks, current, cursor, list, line, |chunks, current| {
        class_slots::emit_class_alloc(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, entry, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, scratch, line);
        emit_node_field(chunks, current, entry, "Value", scratch, line);
        emit_node_field(chunks, current, entry, "Index", cursor, line);
        emit_node_field(chunks, current, entry, "List", list, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, nodes, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });

    emit_index_loop(chunks, current, cursor, nodes, line, |chunks, current| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, nodes, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, entry, line);

        for (field, delta) in [("Next", 1i32), ("Previous", -1i32)] {
            let neighbour = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
            chunks[current].emit_i32_const(delta, line);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, neighbour, line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, neighbour, line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op(Op::I32_GE_S, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, neighbour, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, nodes, line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_op(Op::I32_LT_S, line);
            chunks[current].emit_op(Op::I32_AND, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, nodes, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, neighbour, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_else(line);
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, scratch, line);
            emit_node_field(chunks, current, entry, field, scratch, line);
        }
    });

    nodes
}

/// One node field, in both the .NET spelling and the folded one.
fn emit_node_field(
    chunks: &mut [Chunk],
    current: usize,
    entry: u16,
    field: &str,
    value: u16,
    line: u32,
) {
    for spelling in [field.to_string(), field.to_lowercase()] {
        class_slots::emit_class_set(
            &mut chunks[current],
            ObjSource::Local(entry),
            &field_slot(&spelling),
            ValueSource::Local(value),
            line,
        );
    }
}

/// The node at `index`, or null when the index is outside the list.
fn emit_node_at(chunks: &mut [Chunk], current: usize, list: u16, index: u16, line: u32) {
    let nodes = emit_node_chain(chunks, current, list, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nodes, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nodes, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

/// `list.First` / `list.Last` — the end nodes, null for an empty list.
pub fn emit_linked_list_end(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let list = stash_args(chunks, current, 1, line);
    let index = chunks[current].alloc_scratch(1);
    if last {
        chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
    } else {
        chunks[current].emit_i32_const(0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    emit_node_at(chunks, current, list, index, line);
}

/// `list.Find(v)` / `list.FindLast(v)` — the node holding `v`, or null.
pub fn emit_linked_list_find_node(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let (list, needle) = (base, base + 1);
    let index = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
    if last {
        collections::emit_last_index_of(chunks, current, line);
    } else {
        collections::emit_index_of(chunks, current, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    emit_node_at(chunks, current, list, index, line);
}

/// `list.AddBefore(node, v)` / `list.AddAfter(node, v)` — insert relative to a
/// node's position and answer the node that now holds `v`.
pub fn emit_linked_list_add_relative(chunks: &mut [Chunk], current: usize, after: bool, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let (list, node, value) = (base, base + 1, base + 2);
    let index = chunks[current].alloc_scratch(1);

    // ⛔THE NODE'S STAMPED `Index` IS A SNAPSHOT. Every insertion renumbers the
    // ones after it, so a node taken before an `AddBefore` names the wrong
    // position by the time `AddAfter` reaches it. The node's VALUE still names
    // its place, so the position is looked up afresh.
    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
    emit_get_field(chunks, current, node, "Value", line);
    collections::emit_index_of(chunks, current, line);
    if after {
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_insert_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_node_at(chunks, current, list, index, line);
}

/// `list.RemoveFirst()` / `list.RemoveLast()`.
pub fn emit_linked_list_remove_end(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let list = stash_args(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
    if last {
        collections::emit_pop(chunks, current, line);
    } else {
        collections::emit_shift(chunks, current, line);
    }
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `list.Remove(x)` — `x` is either a NODE or a VALUE, and .NET declares both
/// at arity one. A node carries the `Index` this chain stamped on it; anything
/// else is searched for. Answers the bool the value overload promises.
pub fn emit_linked_list_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let (list, target) = (base, base + 1);
    let index = chunks[current].alloc_scratch(1);

    // A NODE carries the `Value` this chain stamped on it; a bare value does
    // not. Either way the position is the value's, looked up in the list as it
    // stands now.
    let needle = chunks[current].alloc_scratch(1);
    emit_get_field(chunks, current, target, "Value", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, needle, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, needle, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
    collections::emit_index_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_end(line);
}

pub fn emit_vb_collection_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let object = base;
    let items = base + 1;
    let keys = base + 2;

    class_slots::emit_class_alloc(&mut chunks[current], line);
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

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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

pub fn emit_vb_collection_to_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    emit_get_field(chunks, current, base, VB_COLLECTION_ITEMS, line);
    collections::emit_clone(chunks, current, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// The last element, or null on an empty collection — `ConcurrentStack.TryPeek`
/// and `ConcurrentBag.TryPeek`, which both look at the end a `TryTake` would
/// remove.
///
/// Stack: `[collection]` → `[value]`.
pub fn emit_concurrent_peek_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = chunks[current].alloc_scratch(2);
    let len = recv + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    // `len - 1`, spelled as `len + (-1)`: `ops` has no dynamic subtract.
    chunks[current].emit_i32_const(-1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

/// The first element, or null on an empty collection — `ConcurrentQueue.TryPeek`
/// looks at the head a `TryDequeue` would remove.
///
/// Stack: `[collection]` → `[value]`.
pub fn emit_concurrent_peek_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

/// `stack.PushRange(items)` — push every element, in array order, so the LAST
/// element ends up on top.
///
/// Stack: `[stack, items]` → `[null]`.
pub fn emit_concurrent_stack_push_range(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = if argc >= 4 {
        stash_args(chunks, current, 4, line)
    } else {
        stash_args(chunks, current, 2, line)
    };
    let recv = base;
    let items = base + 1;
    let start = chunks[current].alloc_scratch(3);
    let idx = start + 1;
    let len = start + 2;

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "ArgumentNullException",
        "Value cannot be null. (Parameter 'items')",
        line,
    );
    chunks[current].emit_end(line);

    if argc >= 4 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 3, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    }

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `stack.TryPopRange(buffer)` — fill `buffer` from the TOP DOWN and answer how
/// many elements were taken.
///
/// ⛔ TOP TO BOTTOM, like `ToArray`. A stack that was pushed `1,2,3` fills the
/// buffer `3,2` — reading the array forwards would answer `1,2` and reverse the
/// only thing a stack guarantees.
///
/// Stack: `[stack, buffer]` → `[count]`.
pub fn emit_concurrent_stack_try_pop_range(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let base = if argc >= 4 {
        stash_args(chunks, current, 4, line)
    } else {
        stash_args(chunks, current, 2, line)
    };
    let recv = base;
    let buffer = base + 1;
    let start = chunks[current].alloc_scratch(3);
    let idx = start + 1;
    let want = start + 2;

    chunks[current].emit_op_u16(Op::LOCAL_GET, buffer, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_throw_dotnet_exception(
        chunks,
        current,
        "ArgumentNullException",
        "Value cannot be null. (Parameter 'items')",
        line,
    );
    chunks[current].emit_end(line);

    if argc >= 4 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 3, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, want, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, buffer, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, want, line);
    }

    // Take no more than the buffer holds, and no more than the stack has.
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, want, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, want, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, want, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, want, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, buffer, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_pop(chunks, current, line);
    collections::emit_set(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    chunks[current].emit_op_u16(Op::LOCAL_GET, want, line);
}

/// `dict.TryUpdate(key, newValue, comparisonValue)` — a compare-and-set.
/// Answers `False` and changes nothing unless the stored value equals
/// `comparisonValue`.
///
/// Stack: `[dict, key, newValue, comparison]` → `[bool]`.
pub fn emit_dict_try_update(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 4, line);
    let recv = base;
    let key = base + 1;
    let value = base + 2;
    let comparison = base + 3;

    emit_dotnet_dict_key(chunks, current, recv, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparison, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_end(line);
}

/// `dict.TryRemove(key)` — the removed VALUE, or null when the key was absent.
///
/// The `out` half is the walker's: `TryRemove(k, out v)` becomes
/// `(v = TryRemove(k)) != null`, because a tree-registered method carries an
/// arity and no pass-by mode.
///
/// Stack: `[dict, key]` → `[value]`.
pub fn emit_dict_try_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let key = base + 1;
    let value = chunks[current].alloc_scratch(1);

    emit_dotnet_dict_key(chunks, current, recv, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

/// `dict.GetOrAdd(key, value)` — the stored value, INSERTING `value` first when
/// the key is absent. `ecma:map.get` alone answered nothing and wrote nothing.
///
/// Stack: `[dict, key, value]` → `[value]`.
pub fn emit_dict_get_or_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let recv = base;
    let key = base + 1;
    let value_or_factory = base + 2;
    let stored = chunks[current].alloc_scratch(1);

    emit_dotnet_dict_key(chunks, current, recv, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_or_factory, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_or_factory, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::delegates::emit_invoke(chunks, current, 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_or_factory, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stored, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stored, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stored, line);
    chunks[current].emit_end(line);
}

/// `dict.AddOrUpdate(key, addValue, updateFactory)` — store `addValue` when the
/// key is absent, otherwise store `updateFactory(key, existing)`. Answers what
/// was stored.
///
/// Stack: `[dict, key, addValue, updateFactory]` → `[value]`.
pub fn emit_dict_add_or_update(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 4, line);
    let recv = base;
    let key = base + 1;
    let add = base + 2;
    let update = base + 3;
    let stored = chunks[current].alloc_scratch(1);

    emit_dotnet_dict_key(chunks, current, recv, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, update, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    vybe_compiler::primitives::delegates::emit_invoke(chunks, current, 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, add, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stored, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stored, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stored, line);
}

/// `dict.AddOrUpdate(key, addFactory, updateFactory, arg)` — the overload that
/// supplies one factory argument to both delegates.
///
/// Stack: `[dict, key, addFactory, updateFactory, arg]` → `[value]`.
pub fn emit_dict_add_or_update_arg(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 5, line);
    let recv = base;
    let key = base + 1;
    let add_factory = base + 2;
    let update_factory = base + 3;
    let factory_arg = base + 4;
    let stored = chunks[current].alloc_scratch(1);

    emit_dotnet_dict_key(chunks, current, recv, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, update_factory, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, factory_arg, line);
    vybe_compiler::primitives::delegates::emit_invoke(chunks, current, 4, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, add_factory, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, factory_arg, line);
    vybe_compiler::primitives::delegates::emit_invoke(chunks, current, 3, line);

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stored, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stored, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stored, line);
}

/// `HashSet<T>.ToArray()` / `SortedSet<T>.ToArray()`.
///
/// A set is not an array, so the shared array primitives cannot read it
/// directly; `iter_values` is the conversion every ordered read here already
/// goes through.
pub fn emit_set_to_array(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_iter_values(chunks, current, line);
}

/// `ICollection<T>.CopyTo(array[, index])` on a SET receiver — the same shared
/// copy, with the source converted first. `HashSet<T>` and `SortedSet<T>` both
/// declare the one-argument overload, which starts at index 0.
/// Stack: `[set, dest]` or `[set, dest, index]`.
pub fn emit_set_copy_to(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_set_copy_to_ordered(chunks, current, argc, false, line)
}

/// `SortedSet<T>.CopyTo(array[, index])` — the same copy in SORTED order, which
/// is the order the set enumerates in.
pub fn emit_sorted_set_copy_to(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_set_copy_to_ordered(chunks, current, argc, true, line)
}

fn emit_set_copy_to_ordered(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    sorted: bool,
    line: u32,
) {
    let (src, dest, index) = if argc >= 3 {
        let base = stash_args(chunks, current, 3, line);
        (base, base + 1, Some(base + 2))
    } else {
        let base = stash_args(chunks, current, 2, line);
        (base, base + 1, None)
    };
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    if sorted {
        emit_sorted_set_elements(chunks, current, line);
    } else {
        collections::emit_iter_values(chunks, current, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, dest, line);
    match index {
        Some(slot) => chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line),
        None => chunks[current].emit_i32_const(0, line),
    }
    collections::emit_copy_to(chunks, current, line);
}

/// `HashSet<T>.RemoveWhere(predicate)` / `SortedSet<T>.RemoveWhere(predicate)` —
/// drop every element the predicate accepts and answer HOW MANY went.
///
/// The set is drained to an array first: a set cannot be mutated while its own
/// iterator walks it, and the answer is the length of the matched slice.
/// Stack: `[set, predicate]` → `[count]`.
pub fn emit_set_remove_where(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = chunks[current].alloc_scratch(7);
    let pred = recv + 1;
    let values = recv + 2;
    let doomed = recv + 3;
    let index = recv + 4;
    let elem = recv + 5;
    let cursor = recv + 6;

    chunks[current].emit_op_u16(Op::LOCAL_SET, pred, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, values, line);
    let abi = vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    vybe_compiler::primitives::loops::emit_filter(
        chunks, current, abi, pred, values, doomed, index, elem, line,
    );

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, doomed, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, doomed, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    sets::emit_delete(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, doomed, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
}

/// `SortedSet<T>.Reverse()` — the sorted elements, back to front.
pub fn emit_sorted_set_reverse(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_sorted_set_elements(chunks, current, line);
    collections::emit_reverse(chunks, current, line);
}

/// The .NET TRY-PATTERN over a collection: `TryTake(out x)` / `TryPop(out x)` /
/// `TryDequeue(out x)` / `TryPeek(out x)`.
///
/// ⛔It answers a BOOL and writes the item through the out-parameter. A binding
/// that drops the out-parameter and answers the ITEM gets both halves wrong,
/// and silently: `$ok = $bag.TryTake([ref]$v)` then leaves `$v` untouched.
///
/// Stack: `[recv, ref]` → `[bool]`. `remove` distinguishes take from peek,
/// `from_end` the stack end from the queue end.
pub fn emit_try_out(chunks: &mut [Chunk], current: usize, remove: bool, from_end: bool, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let (recv, cell) = (base, base + 1);
    let value = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    match (remove, from_end) {
        (true, true) => collections::emit_pop(chunks, current, line),
        (true, false) => collections::emit_shift(chunks, current, line),
        (false, true) => {
            // The last element, left in place.
            chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            collections::emit_get(chunks, current, line);
        }
        (false, false) => {
            chunks[current].emit_i32_const(0, line);
            collections::emit_get(chunks, current, line);
        }
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cell, line);
    vybe_compiler::primitives::references::emit_cell_store(chunks, current, value, line);
    vybe_compiler::primitives::instructions::core_wasm::bool_const(
        &mut chunks[current],
        line,
        true,
    );
    chunks[current].emit_else(line);
    vybe_compiler::primitives::instructions::core_wasm::bool_const(
        &mut chunks[current],
        line,
        false,
    );
    chunks[current].emit_end(line);
}

/// `TryGetValue(key, out value)` — a BOOL answer with the value written through
/// the out-parameter, the same contract as [`emit_try_out`].
///
/// Stack: `[map, key, ref]` → `[bool]`.
pub fn emit_dict_try_get_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 3, line);
    let (map, key, cell) = (base, base + 1, base + 2);
    let value = chunks[current].alloc_scratch(1);

    emit_dotnet_dict_key(chunks, current, map, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    let has = chunks[current].add_import("ecma:map", "has");
    chunks[current].emit_call(has, 2, line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    let get = chunks[current].add_import("ecma:map", "get");
    chunks[current].emit_call(get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cell, line);
    vybe_compiler::primitives::references::emit_cell_store(chunks, current, value, line);
    vybe_compiler::primitives::instructions::core_wasm::bool_const(
        &mut chunks[current],
        line,
        true,
    );
    chunks[current].emit_else(line);
    // .NET leaves the out-parameter at `default(T)`; zero is what every other
    // miss in this adapter writes.
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cell, line);
    vybe_compiler::primitives::references::emit_cell_store(chunks, current, value, line);
    vybe_compiler::primitives::instructions::core_wasm::bool_const(
        &mut chunks[current],
        line,
        false,
    );
    chunks[current].emit_end(line);
}
