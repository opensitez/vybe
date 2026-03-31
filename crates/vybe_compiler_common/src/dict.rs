//! Dict/map compilation — emits WASM-compatible bytecode using GC struct ops.
//!
//! A dict is an Object with properties (HashMap<String, Value>).
//! These helpers emit the same bytecode pattern every language needs for dict literals,
//! dict access, and dict iteration.

use std::rc::Rc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

/// Emit bytecode to create an empty dict.
/// Stack: [] → [dict_object]
pub fn emit_new(chunk: &mut Chunk, line: u32) {
    // struct_new with 0 properties creates an empty Object
    chunk.emit_op_u16(Op::struct_new, 0, line);
}

/// Emit bytecode to set a key-value pair on a dict.
/// Stack before: [dict, key_value, value_value] (key already on stack)
/// Stack after: [dict] (dict remains for chaining)
///
/// Use `emit_set_const_key` for string literal keys (more common).
pub fn emit_set(chunk: &mut Chunk, line: u32) {
    // struct_set expects [obj, val] with key as constant pool operand.
    // But for dynamic keys, we use the runtime property set path.
    // For now, this is the host-call path for dynamic keys.
    let set_fn = chunk.add_import("vybe:types", "dictAdd");
    chunk.emit_op_u16(Op::call_import, set_fn, line);
    chunk.emit(3, line); // 3 args: dict, key, value
    chunk.emit_op(Op::drop, line); // drop null return, keep dict on stack from before
}

/// Emit bytecode to set a string-literal key on a dict.
/// Stack before: [dict]
/// Stack after: [dict]
pub fn emit_set_const_key(chunk: &mut Chunk, key: &str, line: u32) {
    chunk.emit_op(Op::dup, line); // keep dict on stack
    let key_idx = chunk.add_constant(Value::String(Rc::from(key)));
    chunk.emit_op_u16(Op::struct_set, key_idx, line);
    chunk.emit_op(Op::drop, line); // drop the set result (returns val)
}

/// Emit bytecode to get a value from a dict by string key.
/// Stack before: [dict]
/// Stack after: [value_or_null]
pub fn emit_get_const_key(chunk: &mut Chunk, key: &str, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Rc::from(key)));
    chunk.emit_op_u16(Op::struct_get, key_idx, line);
}

/// Emit bytecode to get a value from a dict by dynamic key (on stack).
/// Stack before: [dict, key]
/// Stack after: [value_or_null]
pub fn emit_get(chunk: &mut Chunk, line: u32) {
    let get_fn = chunk.add_import("vybe:types", "dictItem");
    chunk.emit_op_u16(Op::call_import, get_fn, line);
    chunk.emit(2, line);
}

/// Emit bytecode to get all keys as an array.
/// Stack before: [dict]
/// Stack after: [array_of_keys]
pub fn emit_keys(chunk: &mut Chunk, line: u32) {
    let keys_fn = chunk.add_import("vybe:types", "dictKeys");
    chunk.emit_op_u16(Op::call_import, keys_fn, line);
    chunk.emit(1, line);
}

/// Emit bytecode to get all values as an array.
/// Stack before: [dict]
/// Stack after: [array_of_values]
pub fn emit_values(chunk: &mut Chunk, line: u32) {
    let vals_fn = chunk.add_import("vybe:types", "dictValues");
    chunk.emit_op_u16(Op::call_import, vals_fn, line);
    chunk.emit(1, line);
}

/// Emit bytecode to get all [key, value] pairs as an array.
/// Stack before: [dict]
/// Stack after: [array_of_pairs]
pub fn emit_items(chunk: &mut Chunk, line: u32) {
    let items_fn = chunk.add_import("vybe:array", "dictItems");
    chunk.emit_op_u16(Op::call_import, items_fn, line);
    chunk.emit(1, line);
}

/// Emit bytecode for a complete dict literal with N constant string keys.
/// `entries`: [(key, emit_value_fn)] — caller provides a closure to emit each value.
/// Stack: [] → [dict]
pub fn emit_literal(chunk: &mut Chunk, entry_count: usize, line: u32) -> DictBuilder {
    emit_new(chunk, line);
    DictBuilder { remaining: entry_count }
}

/// Builder for dict literals. Call `entry()` for each key-value pair.
pub struct DictBuilder {
    remaining: usize,
}

impl DictBuilder {
    /// After pushing value onto stack, call this to set it on the dict.
    /// Stack before: [dict, value]
    /// Stack after: [dict]
    pub fn set_entry(&mut self, chunk: &mut Chunk, key: &str, line: u32) {
        let key_idx = chunk.add_constant(Value::String(Rc::from(key)));
        chunk.emit_op_u16(Op::struct_set, key_idx, line);
        chunk.emit_op(Op::drop, line);
        self.remaining -= 1;
    }
}
