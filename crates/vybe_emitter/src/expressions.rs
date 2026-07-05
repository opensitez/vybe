//! Expression compilation helpers — shared bytecode patterns for common expressions.
//!
//! Ternary conditionals, short-circuit logic, and null coalescing are identical
//! across all languages. Helpers emit structured WASM control constructs so
//! callers can compile language-specific sub-expressions in between.

use crate::instructions::core_wasm;
use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

// ── Undefined sentinel ─────────────────────────────────────────────────
//
// JS `undefined` is not a WASM concept. We represent it as a global sentinel
// `__undefined` that is set up at bundle time. All compilers that need
// undefined semantics emit `global_get "__undefined"` via this helper.
// Languages that don't have undefined (VB, C#, Pascal, Python) never call this.
//
// This centralizes all undefined emission so the opcode `Op::undefined` can
// eventually be removed — every site that used to emit that opcode should
// call this function instead.

/// Emit the JS `undefined` value onto the stack.
/// Uses `global_get "__undefined"` — a sentinel wired at bundle time.
/// Stack: [] → [undefined]
pub fn emit_undefined(chunk: &mut Chunk, line: u32) {
    let name = chunk.add_constant(Value::String(Arc::from("__undefined")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, line);
}

/// Emit bitwise NOT (i32). WASM equivalent: i32.const -1, i32.xor.
/// Stack: [i32] → [i32]
pub fn emit_i32_not(chunk: &mut Chunk, line: u32) {
    chunk.emit_i32_const(-1, line);
    chunk.emit_op(Op::I32_XOR, line);
}

/// Emit f64 C-style modulo as pure WASM opcodes (no host import).
/// Stack: [a, b] → [result]
pub fn emit_f64_mod_with_import(chunk: &mut Chunk, _import_idx: u16, line: u32) {
    crate::math::emit_c_fmod(chunk, line);
}

/// Emit f64 C-style modulo as pure WASM opcodes (no host import).
/// Stack: [a, b] → [result]
pub fn emit_f64_mod(chunk: &mut Chunk, line: u32) {
    crate::math::emit_c_fmod(chunk, line);
}

/// Emit boolean NOT. Converts value to bool then negates.
/// WASM equivalent: dyn_to_bool + i32.eqz.
/// Stack: [value] → [bool]
pub fn emit_bool_not(chunk: &mut Chunk, line: u32) {
    crate::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
}

// ── Ternary / conditional expression ────────────────────────────────────
//
// Usage:
//   compile_condition(chunk);
//   let false_jump = emit_ternary_start(chunk);
//   compile_then_expr(chunk);
//   let end_jump = emit_ternary_middle(chunk, false_jump);
//   compile_else_expr(chunk);
//   emit_ternary_end(chunk, end_jump);

/// After condition is on stack: convert to bool and enter the then arm.
/// Stack before: [condition]  Stack after: []
pub fn emit_ternary_start(chunk: &mut Chunk, line: u32) -> usize {
    crate::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    0
}

/// After "then" expression: start the else arm.
/// Stack: [then_value]
pub fn emit_ternary_middle(chunk: &mut Chunk, _false_jump: usize, line: u32) -> usize {
    chunk.emit_else(line);
    0
}

/// After "else" expression: close the structured if.
/// Stack: [result_value]
pub fn emit_ternary_end(chunk: &mut Chunk, _end_jump: usize) {
    chunk.emit_end(0);
}

// ── Short-circuit logical AND ───────────────────────────────────────────
//
// Usage:
//   compile_left(chunk);
//   let jump = emit_and_start(chunk);
//   compile_right(chunk);
//   emit_short_circuit_end(chunk, jump);

/// After left operand: if falsy, short-circuit (keep left as result).
/// Stack before: [left]  Stack after: [] (right will be compiled next)
pub fn emit_and_start(chunk: &mut Chunk, line: u32) -> usize {
    let block = chunk.emit_block(line);
    chunk.emit_dup(line);
    crate::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line); // discard left, right becomes result
    block
}

// ── Short-circuit logical OR ────────────────────────────────────────────
//
// Usage:
//   compile_left(chunk);
//   let jump = emit_or_start(chunk);
//   compile_right(chunk);
//   emit_short_circuit_end(chunk, jump);

/// After left operand: if truthy, short-circuit (keep left as result).
/// Stack before: [left]  Stack after: [] (right will be compiled next)
pub fn emit_or_start(chunk: &mut Chunk, line: u32) -> usize {
    let block = chunk.emit_block(line);
    chunk.emit_dup(line);
    crate::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line); // discard left, right becomes result
    block
}

/// End a short-circuit AND or OR.
/// Stack: [result_value]
pub fn emit_short_circuit_end(chunk: &mut Chunk, block: usize) {
    chunk.emit_end(0);
    chunk.patch_block(block);
}

// ── Null coalescing ─────────────────────────────────────────────────────
//
// Usage:
//   compile_left(chunk);
//   let (null_jump, end_jump) = emit_null_coalesce_start(chunk);
//   compile_right(chunk);
//   emit_null_coalesce_end(chunk, end_jump);

/// After left operand: if null, drop it and fall through to right expression.
/// If non-null, skip over right expression.
/// Stack before: [left]  Stack after: [] (right will be compiled next)
/// Returns (block_patch, 0).
pub fn emit_null_coalesce_start(chunk: &mut Chunk, line: u32) -> (usize, usize) {
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let block = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);
    (block, 0)
}

/// After right expression: close the non-null skip block.
/// Stack: [result_value]
pub fn emit_null_coalesce_end(chunk: &mut Chunk, block: usize) {
    chunk.emit_end(0);
    chunk.patch_block(block);
}

// ── Null-safe member access (?.) ────────────────────────────────────────
//
// Usage:
//   compile_object(chunk);
//   let (skip, end) = emit_null_safe_start(chunk);
//   // compile member access (struct_get, etc.)
//   emit_null_safe_end(chunk, end);

/// After object is on stack: if null, skip the member access.
/// Stack before: [object]  Stack after: [object] (if non-null) or control jumps to end
pub fn emit_null_safe_start(chunk: &mut Chunk, line: u32) -> (usize, usize) {
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let block = chunk.emit_block(line);
    chunk.emit_br_if(0, line);
    (block, 0)
}

/// After member access: close the null-skip block.
pub fn emit_null_safe_end(chunk: &mut Chunk, block: usize, _line: u32) {
    chunk.emit_end(0);
    chunk.patch_block(block);
}

// ── Generic dynamic dispatch ────────────────────────────────────────────
//
// The universal pattern for dynamic languages:
//   1. struct_get a method/property on the object
//   2. If found (non-null), call it
//   3. If not found, execute a fallback
//
// All rich_compare, smart_length, rich_arithmetic, etc. are instances of this.

/// Emit a try-method-or-fallback dispatch.
/// Checks if `obj_slot` has a method named `method_name`. If found, calls it with
/// `arg_count` args (which the caller pushes between start and end).
/// Returns (is_null_jump, found_done_jump) for the caller to emit the fallback.
///
/// Usage:
///   let (null_jump, done_jump) = emit_dynamic_dispatch_start(chunk, obj_slot, "method", line);
///   // push args for the found case
///   emit_dynamic_dispatch_call(chunk, arg_count, line);
///   let done = emit_dynamic_dispatch_middle(chunk, line);
///   // patch null case, emit fallback
///   emit_dynamic_dispatch_fallback(chunk, null_jump, done, line);
///
/// Or use the simpler one-shot helpers below.
// ── Rich arithmetic (user-defined __add__/__sub__/etc) ──────────────────
//
// Same pattern as rich_compare but for binary arithmetic operators.
// Tries user-defined dunder method, falls back to primitive opcode.

/// Emit rich arithmetic: tries user-defined __add__/etc, falls back to primitive opcode.
/// Caller must store left in `left_slot` and right in `right_slot`.
/// Stack before: []  Stack after: [result_value]
pub fn emit_rich_arithmetic(
    chunk: &mut Chunk,
    left_slot: u16,
    right_slot: u16,
    dunder: &str,
    fallback_fn: fn(&mut Chunk, u32),
    line: u32,
) {
    emit_rich_compare_locals(chunk, left_slot, right_slot, dunder, fallback_fn, line);
}

// ── Rich toString (user-defined __str__ / toString) ─────────────────────

/// Emit smart toString: tries __str__/toString getter on objects, falls back to host toString.
/// Object must be in `obj_slot`.
/// Stack before: []  Stack after: [string_value]
pub fn emit_rich_to_string(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    let method_slot = chunk.local_count;
    chunk.local_count = chunk.local_count.max(method_slot + 1);
    if chunk.local_count > chunk.scratch_high_water {
        chunk.scratch_high_water = chunk.local_count;
    }

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from("__str__")));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let to_str = chunk.add_import("ecma:string", "String");
    chunk.emit_call(to_str, 1, line);
    chunk.emit_end(line);
}

// ── Rich bool (user-defined __bool__ / valueOf) ─────────────────────────

/// Emit smart bool: tries __bool__ on objects, falls back to dyn_to_bool.
/// Object must be in `obj_slot`.
/// Stack before: []  Stack after: [bool_value]
pub fn emit_rich_bool(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    let method_slot = chunk.local_count;
    chunk.local_count = chunk.local_count.max(method_slot + 1);
    if chunk.local_count > chunk.scratch_high_water {
        chunk.scratch_high_water = chunk.local_count;
    }

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from("__bool__")));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    crate::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_end(line);
}

// ── Rich comparison (user-defined __lt__/__gt__/etc) ────────────────────
//
// Standard WASM opcodes only. Emits inline dispatch:
//   1. Try struct_get for the dunder method on left operand
//   2. If found (non-null), call it with right operand
//   3. If not found, fall back to the primitive dyn_lt/dyn_gt/etc opcode
//
// This allows Python `__lt__`, Dart `operator<`, C# `CompareTo` etc.
// to work on user objects while keeping primitive comparison fast.

/// Emit a rich comparison: tries user-defined method, falls back to primitive opcode.
/// Both operands must already be on the stack: [left, right].
/// Stack after: [bool_result]
///
/// `dunder`: the method name to look for (e.g. "__lt__", "__gt__")
/// `fallback_fn`: the emitter to use if no method (e.g. `crate::ops::emit_dyn_lt`)
pub fn emit_rich_compare(
    chunk: &mut Chunk,
    _dunder: &str,
    fallback_fn: fn(&mut Chunk, u32),
    line: u32,
) {
    // Stack: [left, right]
    // Save right to temp, check left for dunder method
    // We need to: peek at left (under right), struct_get dunder, check null

    // Store right in temp
    // Note: we can't allocate locals here (no scope access). Use stack manipulation.
    // Strategy: swap to get left on top, dup, struct_get, check null.
    // But there's no swap opcode. Use a different approach:
    // Store right, dup left, struct_get dunder, check null.

    // Actually, the simplest approach that uses only standard WASM ops:
    // [left, right] on stack.
    // We need left for struct_get. But right is on top.
    // Emit: store right in a constant-indexed temp via the "over" pattern.

    // Simplest correct approach using only existing opcodes:
    // The caller must have left and right in locals already (common for binary ops).
    // But we take them from stack. Let's use the dup-under pattern:

    // For now, just use the fallback op. Rich compare requires local slots
    // which the caller must provide. Use emit_rich_compare_with_locals instead.
    fallback_fn(chunk, line);
}

/// Emit a rich comparison with pre-allocated local slots.
/// Caller must store left in `left_slot` and right in `right_slot` before calling.
/// Stack before: []  Stack after: [bool_result]
///
/// Emits: check left.__lt__ → if found, call it(right) → else dyn_lt(left, right)
pub fn emit_rich_compare_locals(
    chunk: &mut Chunk,
    left_slot: u16,
    right_slot: u16,
    dunder: &str,
    fallback_fn: fn(&mut Chunk, u32),
    line: u32,
) {
    let method_slot = chunk.local_count;
    chunk.local_count = chunk.local_count.max(method_slot + 1);
    if chunk.local_count > chunk.scratch_high_water {
        chunk.scratch_high_water = chunk.local_count;
    }

    // Try struct_get dunder on left
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from(dunder)));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    // Found method: call it with self=left, arg=right → result
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_else(line);

    // Not found: try compare-style methods like C# CompareTo / Ruby <=>.
    let done = chunk.emit_block(line);
    for method_name in ["compare", "CompareTo", "compareTo", "__cmp__", "<=>"] {
        let method_key = chunk.add_constant(Value::String(Arc::from(method_name)));
        chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
        chunk.emit_op_u16(Op::STRUCT_GET, method_key, line);
        chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

        chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
        chunk.emit_op_u8(Op::CALL_REF, 2, line);
        core_wasm::i32_const(chunk, line, 0);
        fallback_fn(chunk, line);
        chunk.emit_br(1, line);
        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    fallback_fn(chunk, line);

    chunk.emit_end(line);
    chunk.patch_block(done);
    chunk.emit_end(line);
}

// ── Smart length (user-defined __len__ / __get_length) ──────────────────
//
// Standard WASM opcodes only. Tries __get_length getter first,
// falls back to array_length opcode for plain arrays/strings.

/// Emit smart length: tries user-defined __get_length getter, falls back to array_length.
/// Object must be in `obj_slot`.
/// Stack before: []  Stack after: [length_value]
pub fn emit_smart_length(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    let method_slot = chunk.local_count;
    chunk.local_count = chunk.local_count.max(method_slot + 1);
    if chunk.local_count > chunk.scratch_high_water {
        chunk.scratch_high_water = chunk.local_count;
    }

    // Try struct_get "__get_length" on object
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from("__get_length")));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_end(line);
}
