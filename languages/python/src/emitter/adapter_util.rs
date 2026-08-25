//! Shared emit helpers for the Python stdlib adapters.
//!
//! Every adapter in this directory needs the same handful of moves — pop the
//! call's arguments into scratch slots, read one back, set a property, build a
//! fresh object, call a host function. The older files each carry their own
//! copy; the modules added since share these.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

pub fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

pub fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Pop `argc` call arguments into `argc` consecutive scratch slots, leftmost
/// first, and return the base slot. The stack is popped right-to-left, which
/// is why the loop counts down.
pub fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc.max(1) as u16);
    for offset in (0..argc as u16).rev() {
        lset(&mut chunks[current], base + offset, line);
    }
    base
}

/// Pop exactly `want` arguments into slots, padding a short call with
/// `undefined` and dropping a long one's extras, so the emitted code can read
/// a fixed slot layout regardless of how the call was written.
pub fn stash_exact(chunks: &mut [Chunk], current: usize, argc: u8, want: u16, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(want.max(1));
    for offset in (want..argc as u16).rev() {
        let _ = offset;
        chunks[current].emit_op(Op::DROP, line);
    }
    for offset in (0..want).rev() {
        if offset < argc as u16 {
            lset(&mut chunks[current], base + offset, line);
        } else {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            lset(&mut chunks[current], base + offset, line);
        }
    }
    base
}

pub fn string_key(chunk: &mut Chunk, key: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(key)))
}

/// `<obj> <value> → ` — sets `obj.key = value`, consuming both.
pub fn struct_set(chunk: &mut Chunk, key: &str, line: u32) {
    let k = string_key(chunk, key);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}

/// `<obj> → <value>`.
pub fn struct_get(chunk: &mut Chunk, key: &str, line: u32) {
    let k = string_key(chunk, key);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

pub fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

/// A fresh empty object left on the stack.
pub fn new_object(chunk: &mut Chunk, line: u32) {
    chunk.emit_struct_new(0, 0, line);
}

/// Build an object with `__type` set and each `(key, slot)` copied in, left on
/// the stack. Used by every adapter that returns a tagged record.
pub fn new_tagged(chunk: &mut Chunk, type_name: &str, fields: &[(&str, u16)], line: u32) {
    new_object(chunk, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(type_name, line);
    struct_set(chunk, "__type", line);
    for (key, slot) in fields {
        chunk.emit_dup(line);
        lget(chunk, *slot, line);
        struct_set(chunk, key, line);
    }
}

/// Attach `chunk_index` as the object's Call slot, so `obj(...)` invokes it.
/// The shared call path (`primitives/calls.rs`, gated on the profile's
/// `callable_objects`) probes exactly this key and passes the object itself
/// as the leading argument — which is why such a helper chunk's arity counts
/// the receiver.
pub fn set_call_slot(chunk: &mut Chunk, chunk_index: usize, line: u32) {
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::REF_FUNC, chunk_index as u16, line);
    chunk.emit(0, line);
    let key = vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Call);
    let k = string_key(chunk, &key);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}
