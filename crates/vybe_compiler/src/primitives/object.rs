//! Common object helpers shared by language adapters.
//!
//! These are ECMA-shaped object/value operations. Language frontends should
//! normalize their own API names (`java.util.Objects.equals`, PHP object
//! helpers, etc.) into profile builtins that route here when the semantics are
//! genuinely shared.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

/// Dynamic value equality. Stack: [left, right] -> [Bool]
pub fn emit_equals(chunk: &mut Chunk, line: u32) {
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// Null test. Stack: [value] -> [Bool]
pub fn emit_is_null(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::REF_IS_NULL, line);
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// Non-null test. Stack: [value] -> [Bool]
pub fn emit_non_null(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// Set an object monitor's notified marker.
///
/// This is the object side of monitor notify/notifyAll. Languages still own
/// their scheduling/catch semantics, but the object-state mutation is common.
/// Stack: [object] -> [null]
pub fn emit_monitor_notify(chunk: &mut Chunk, line: u32) {
    chunk.emit_bool_const(true, line);
    let key = chunk.add_constant(Value::String(Arc::from("__j_notified")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Deterministic object hash based on ECMA string conversion.
///
/// This mirrors Java's common `31*h + codeUnit` polynomial and is useful for
/// language APIs that need stable object/value hashing without identity hash
/// support. Null hashes to 0.
/// Stack: [value] -> [i32]
pub fn emit_hash_code(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(4);
    let text = value + 1;
    let hash = value + 2;
    let index = value + 3;
    let to_string = chunk.add_import("ecma:string", "String");
    let length = chunk.add_import("wasm:js-string", "length");
    let char_code_at = chunk.add_import("wasm:js-string", "charCodeAt");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(to_string, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hash, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);

    let outer = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    chunk.emit_i32_const(31, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hash, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_call(char_code_at, 2, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hash, line);

    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, hash, line);
    chunk.emit_end(line);
}

/// Hash an array of values using the Java/List-style accumulator.
/// Stack: [array] -> [i32]
pub fn emit_hash_array(chunk: &mut Chunk, line: u32) {
    let items = chunk.alloc_scratch(3);
    let hash = items + 1;
    let index = items + 2;
    let length = chunk.add_import("ecma:array", "length");
    let get = chunk.add_import("ecma:array", "get");

    chunk.emit_op_u16(Op::LOCAL_SET, items, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hash, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);

    let outer = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    chunk.emit_i32_const(31, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hash, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_call(get, 2, line);
    emit_hash_code(chunk, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hash, line);

    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, hash, line);
}

/// Null-aware comparator dispatch. Stack: [a, b, comparator] -> [value]
pub fn emit_compare(chunk: &mut Chunk, line: u32) {
    let cmp = chunk.alloc_scratch(3);
    let b = cmp + 1;
    let a = cmp + 2;

    chunk.emit_op_u16(Op::LOCAL_SET, cmp, line);
    chunk.emit_op_u16(Op::LOCAL_SET, b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a, line);

    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, cmp, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 2, 1, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// String conversion with a null fallback. Stack: [value, fallback] -> [string]
pub fn emit_to_string_or(chunk: &mut Chunk, line: u32) {
    let fallback = chunk.alloc_scratch(2);
    let value = fallback + 1;
    let to_string = chunk.add_import("ecma:string", "String");

    chunk.emit_op_u16(Op::LOCAL_SET, fallback, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, fallback, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_string_const("null", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, fallback, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(to_string, 1, line);
    chunk.emit_end(line);
}

// ── Method binding on an object ─────────────────────────────────────────
//
// Object-member binding primitives, needed by crates BELOW `vybe_compiler`
// (`platforms/dotnet` guid/uri adapters), so they cannot live with the class
// model. Binding a function onto an object is object mutation, not class
// policy — this is their home.
//
// The spelling-guess table that used to live here (flexclassplan.md §1b,
// source #6) is GONE. Cross-language reach is a PROTOCOL SLOT — one numeric
// key per role — not a list of every language's name for the same idea.

/// Bind an instance method on the object: this.<method_name> = ref_func(chunk_idx).
/// Emits: local_get this → ref_func ci → struct_set key → drop
/// Stack: unchanged
pub fn emit_bind_method(
    chunk: &mut Chunk,
    this_slot: u16,
    method_name: &str,
    method_chunk_idx: usize,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, method_chunk_idx as u16, line);
    chunk.emit(0, line); // 0 upvalues (upvalue capture is compiler-specific)
    let key = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

pub fn emit_bind_bound_method(
    chunk: &mut Chunk,
    this_slot: u16,
    method_name: &str,
    method_chunk_idx: usize,
    rest_fixed_count: Option<u8>,
    distinct_per_instance: bool,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, method_chunk_idx as u16, line);
    if distinct_per_instance {
        // Capture the receiver as upvalue[0] so this funcref is a FRESH object
        // for every instance (uv_count > 0 means `REF_FUNC` does NOT intern it).
        // The upvalue is inert for dispatch — `self` still routes through the
        // `__vybe_method_receiver` property below — but it gives each instance's
        // bound method a distinct identity (`C().f is C().f` is False under
        // `methods_bind_on_access`; `c.f is c.f` stays True since the binding is
        // stored once per instance).
        chunk.emit(1, line); // 1 upvalue
        chunk.emit(1, line); // is_local = true
        chunk.emit((this_slot >> 8) as u8, line); // capture index (u16, big-endian)
        chunk.emit((this_slot & 0xff) as u8, line);
    } else {
        chunk.emit(0, line); // 0 upvalues (bind-at-call languages share the fn)
    }
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let receiver_key = chunk.add_constant(Value::String(Arc::from("__vybe_method_receiver")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, receiver_key, line);
    if let Some(fixed_count) = rest_fixed_count {
        emit_stamp_rest_metadata(chunk, fixed_count, line);
    }
    let key = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

pub fn emit_stamp_rest_metadata(chunk: &mut Chunk, fixed_count: u8, line: u32) {
    chunk.emit_dup(line);
    chunk.emit_f64_const(fixed_count as f64, line);
    let key = chunk.add_constant(Value::String(Arc::from("__vybe_rest_fixed_arity")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

/// Bind a method under its own name AND under the numeric key of the PROTOCOL
/// SLOT it fills, so a caller in any language reaches it through the role
/// rather than by guessing a spelling.
///
/// This replaces a synonym table that bound every cross-language spelling of a
/// special method as its own property — a Python `__str__` also published
/// `toString`, `tostring`, `ToString`, `to_s` and `__toString`. Five extra
/// names in the namespace user members live in is what let a synonym set
/// capture an unrelated user method. A slot key is derived from a NUMBER, so
/// nothing a user can type collides with it.
///
/// `slot` is `None` for ordinary members, which then bind under one name only.
/// Stack: unchanged
pub fn emit_bind_method_with_slot(
    chunk: &mut Chunk,
    this_slot: u16,
    method_name: &str,
    slot: Option<vybe_ast::ProtocolSlot>,
    method_chunk_idx: usize,
    rest_fixed_count: Option<u8>,
    line: u32,
) {
    let mut bind = |name: &str| {
        emit_bind_method(chunk, this_slot, name, method_chunk_idx, line);
        if let Some(fixed_count) = rest_fixed_count {
            chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
            let key = chunk.add_constant(Value::String(Arc::from(name)));
            chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
            emit_stamp_rest_metadata(chunk, fixed_count, line);
            chunk.emit_op(Op::DROP, line);
        }
    };
    bind(method_name);
    if let Some(slot) = slot {
        bind(&vybe_ast::protocol_slot_key(slot));
    }
}

/// Stamp shared reflection/class metadata on an existing object when the class
/// name is known as a runtime value. This is useful for dynamic languages that
/// normalize class-like prototype objects but cannot bake a static class token
/// into the emitter call.
pub fn emit_retype_object_dynamic(
    chunks: &mut [Chunk],
    current: usize,
    this_slot: u16,
    class_name_slot: u16,
    line: u32,
) {
    stamp_local_dynamic_field(
        &mut chunks[current],
        this_slot,
        crate::primitives::reflection::FIELD_TYPE,
        class_name_slot,
        line,
    );
    stamp_local_dynamic_field(
        &mut chunks[current],
        this_slot,
        crate::primitives::reflection::FIELD_TYPE_NAME,
        class_name_slot,
        line,
    );
    stamp_local_string_field(
        &mut chunks[current],
        this_slot,
        crate::primitives::reflection::FIELD_KIND,
        crate::primitives::reflection::ReflectKind::Object.as_str(),
        line,
    );
    stamp_local_dynamic_field(
        &mut chunks[current],
        this_slot,
        "__control_name",
        class_name_slot,
        line,
    );

    let types_key = chunks[current].add_constant(Value::String(Arc::from(
        crate::primitives::reflection::FIELD_TYPES,
    )));
    chunks[current].emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, types_key, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    let init_block = chunks[current].emit_block(line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(0, line);
    chunks[current].emit_op(Op::DROP, line);
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(init_block);

    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, class_name_slot, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, types_key, line);
}
/// Bind a property getter as __get_<name> on the instance.
/// The getter_chunk_idx should point to a compiled closure with arity=1 (self/this).
/// Stack: unchanged
pub fn emit_bind_getter(
    chunk: &mut Chunk,
    this_slot: u16,
    prop_name: &str,
    getter_chunk_idx: usize,
    line: u32,
) {
    let get_name = format!("__get_{}", prop_name);
    emit_bind_method(chunk, this_slot, &get_name, getter_chunk_idx, line);
}
/// Bind a property setter as __set_<name> on the instance.
/// The setter_chunk_idx should point to a compiled closure with arity=2 (self/this, value).
/// Stack: unchanged
pub fn emit_bind_setter(
    chunk: &mut Chunk,
    this_slot: u16,
    prop_name: &str,
    setter_chunk_idx: usize,
    line: u32,
) {
    let set_name = format!("__set_{}", prop_name);
    emit_bind_method(chunk, this_slot, &set_name, setter_chunk_idx, line);
}

pub fn stamp_local_dynamic_field(
    chunk: &mut Chunk,
    slot: u16,
    field: &str,
    value_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from(field)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}
pub fn stamp_local_string_field(chunk: &mut Chunk, slot: u16, field: &str, value: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    set_string_field(chunk, field, value, line);
}
pub fn set_string_field(chunk: &mut Chunk, field: &str, value: &str, line: u32) {
    chunk.emit_string_const(value, line);
    let key = chunk.add_constant(Value::String(Arc::from(field)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}
