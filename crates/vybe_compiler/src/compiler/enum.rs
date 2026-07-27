//! Shared enum machinery — the TS-shaped bidirectional enum object.
//!
//! An enum compiles to a single object carrying BOTH directions:
//! `{ Red: 0, Green: 1, "0": "Red", "1": "Green" }` — forward `name → value`
//! (values stay bare ints, so flags/arithmetic/comparison/casts never break)
//! plus a reverse `value → name` map. These helpers implement the enum
//! operations as generic RUNTIME reads on that object, so no language has to
//! hand-roll compile-time ordinal tables. Any language whose enums use this
//! shape (C#, VB, …) shares this one emitter.
//!
//! Reads use `ecma:object.get` (a raw property-bag read) rather than an index
//! expression: an enum object carries an index getter that does array-position
//! lookup, which only matches sequential values — the raw read hits the
//! reverse field directly.

use vybe_bytecode::{Chunk, Op};

use crate::compiler::instructions::host;

/// `value → name` (enum `ToString`, `Enum.GetName`, `Enum.Format("G")`).
/// Stack: `[enumObj, value]` → `[string]`.
///
/// Only NUMERIC values map through the reverse field; a value that is already a
/// name string (e.g. an `Enum.Parse` result flowing into `ToString`) passes
/// through `String()` unchanged. Numeric values that aren't defined members
/// fall back to `String(value)` (matches .NET's numeric `ToString`).
pub fn emit_value_to_name(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    let obj = chunk.alloc_scratch(1);
    let name = chunk.alloc_scratch(1);
    // Stack pushed as [enumObj, value]; pop value first.
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    // typeof(value) === "number"
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("number", line);
    crate::compiler::ops::emit_dyn_eq(chunk, line);
    crate::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    // name = ecma:object.get(enumObj, "" + value)  (raw reverse-field read)
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    crate::compiler::ops::emit_dyn_add(chunk, line);
    host::emit(chunk, "ecma:object", "get", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name, line);

    // name undefined ? String(value) : name
    chunk.emit_op_u16(Op::LOCAL_GET, name, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(chunk, "ecma:string", "String", 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, name, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    // Non-numeric (already a name string): pass through unchanged.
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(chunk, "ecma:string", "String", 1, line);
    chunk.emit_end(line);
}

/// Case-sensitive `name → validated name or null` (enum `Parse` / `IsDefined` /
/// `TryParse`). Stack: `[enumObj, input]` → `[string | null]`.
///
/// `input` names a member iff a raw read of the enum object yields its NUMERIC
/// forward field. A numeric-string input would instead hit a reverse
/// (value→name) field — a string — and is correctly rejected. Returns the
/// input (== the canonical name on an exact match) or null.
pub fn emit_name_to_member_or_null(chunk: &mut Chunk, line: u32) {
    let input = chunk.alloc_scratch(1);
    let obj = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, input, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);
    // Coerce input to a string once.
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    host::emit(chunk, "ecma:string", "String", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input, line);

    // typeof(ecma:object.get(enumObj, input)) === "number" ? input : null
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    host::emit(chunk, "ecma:object", "get", 2, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("number", line);
    crate::compiler::ops::emit_dyn_eq(chunk, line);
    crate::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_else(line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_end(line);
}

/// `HasFlag` — `(value & flag) === flag`. Stack: `[value, flag]` → `[bool]`.
pub fn emit_has_flag(chunk: &mut Chunk, line: u32) {
    let flag = chunk.alloc_scratch(1);
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, flag, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, flag, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, flag, line);
    crate::compiler::ops::emit_dyn_eq(chunk, line);
}
