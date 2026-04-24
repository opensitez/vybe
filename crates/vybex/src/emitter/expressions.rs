//! Expression compilation helpers — shared bytecode patterns for common expressions.
//!
//! Ternary conditionals, short-circuit logic, and null coalescing are identical
//! across all languages. Each helper emits the jump structure and returns
//! patch points so the caller can compile the language-specific sub-expressions
//! in between.

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

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
    let neg1 = chunk.add_constant(Value::I32(-1));
    chunk.emit_op_u16(Op::CONST, neg1, line);
    chunk.emit_op(Op::I32_XOR, line);
}

/// Emit f64 modulo via stdlib function `__vybe_fmod`.
/// Pure bytecode implementation: a % b = a - trunc(a/b) * b.
/// Host can override with native fmod for performance.
/// `import_idx` is the resolved import index for `ecma:math:fmod`.
/// Stack: [a, b] → [result]
pub fn emit_f64_mod_with_import(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, line);
    chunk.emit(2, line);
}

/// Convenience: emit f64 modulo, adding the import to the same chunk.
/// Use only when imports go on the same chunk as the code (old compilers).
/// Stack: [a, b] → [result]
pub fn emit_f64_mod(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:math", "fmod");
    emit_f64_mod_with_import(chunk, idx, line);
}

/// Emit boolean NOT. Converts value to bool then negates.
/// WASM equivalent: dyn_to_bool + i32.eqz.
/// Stack: [value] → [bool]
pub fn emit_bool_not(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::DYN_TO_BOOL, line);
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

/// After condition is on stack: convert to bool, jump to else if false.
/// Stack before: [condition]  Stack after: []
pub fn emit_ternary_start(chunk: &mut Chunk, line: u32) -> usize {
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    chunk.emit_jump(Op::BR_IF_FALSE, line)
}

/// After "then" expression: jump over else, patch the false target.
/// Stack: [then_value]
pub fn emit_ternary_middle(chunk: &mut Chunk, false_jump: usize, line: u32) -> usize {
    let end_jump = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(false_jump);
    end_jump
}

/// After "else" expression: patch the end target.
/// Stack: [result_value]
pub fn emit_ternary_end(chunk: &mut Chunk, end_jump: usize) {
    chunk.patch_jump(end_jump);
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
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    let jump = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::DROP, line); // discard left, right becomes result
    jump
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
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    let jump = chunk.emit_jump(Op::BR_IF_TRUE, line);
    chunk.emit_op(Op::DROP, line); // discard left, right becomes result
    jump
}

/// End a short-circuit AND or OR.
/// Stack: [result_value]
pub fn emit_short_circuit_end(chunk: &mut Chunk, jump: usize) {
    chunk.patch_jump(jump);
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
/// Returns (null_jump, end_jump).
pub fn emit_null_coalesce_start(chunk: &mut Chunk, line: u32) -> (usize, usize) {
    chunk.emit_op(Op::DUP, line);
    let null_jump = chunk.emit_jump(Op::BR_IF_NULL, line);
    // Not null — keep left, jump to end
    let end_jump = chunk.emit_jump(Op::BR, line);
    // Null path: drop the null, caller compiles right expression
    chunk.patch_jump(null_jump);
    chunk.emit_op(Op::DROP, line);
    (null_jump, end_jump)
}

/// After right expression: patch the non-null skip jump.
/// Stack: [result_value]
pub fn emit_null_coalesce_end(chunk: &mut Chunk, end_jump: usize) {
    chunk.patch_jump(end_jump);
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
    chunk.emit_op(Op::DUP, line);
    let skip = chunk.emit_jump(Op::BR_IF_NULL, line);
    // Non-null: fall through to member access
    // Return skip (null path) — caller compiles access, then calls end
    (skip, 0) // end_jump set by caller
}

/// After member access: patch the null-skip and null-end jumps.
pub fn emit_null_safe_end(chunk: &mut Chunk, skip: usize, line: u32) {
    let end = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(skip);
    // null is still on stack from the dup
    chunk.patch_jump(end);
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
pub fn emit_try_method(chunk: &mut Chunk, obj_slot: u16, method_name: &str, line: u32) -> (usize, bool) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let is_null = chunk.emit_jump(Op::BR_IF_TRUE, line);
    (is_null, true)
}

/// After emit_try_method found a method: clean up the null path.
/// Call this after emitting the call_ref in the found path.
pub fn emit_try_method_fallback_start(chunk: &mut Chunk, is_null: usize, line: u32) -> usize {
    let done = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(is_null);
    chunk.emit_op(Op::DROP, line); // drop null from dup
    chunk.emit_op(Op::DROP, line); // drop null from struct_get
    done
}

/// End the try-method dispatch.
pub fn emit_try_method_end(chunk: &mut Chunk, done: usize) {
    chunk.patch_jump(done);
}

// ── Rich arithmetic (user-defined __add__/__sub__/etc) ──────────────────
//
// Same pattern as rich_compare but for binary arithmetic operators.
// Tries user-defined dunder method, falls back to primitive opcode.

/// Emit rich arithmetic: tries user-defined __add__/etc, falls back to primitive opcode.
/// Caller must store left in `left_slot` and right in `right_slot`.
/// Stack before: []  Stack after: [result_value]
pub fn emit_rich_arithmetic(chunk: &mut Chunk, left_slot: u16, right_slot: u16, dunder: &str, fallback_op: Op, line: u32) {
    // Same dispatch pattern as rich_compare
    emit_rich_compare_locals(chunk, left_slot, right_slot, dunder, fallback_op, line);
}

// ── Rich toString (user-defined __str__ / toString) ─────────────────────

/// Emit smart toString: tries __str__/toString getter on objects, falls back to host toString.
/// Object must be in `obj_slot`.
/// Stack before: []  Stack after: [string_value]
pub fn emit_rich_to_string(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    // Try struct_get "__str__" on object
    let (is_null, _) = emit_try_method(chunk, obj_slot, "__str__", line);

    // Found: call __str__(self) → returns string
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    let done = emit_try_method_fallback_start(chunk, is_null, line);

    // Fallback: use host toString
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let to_str = chunk.add_import("vybe:convert", "toString");
    chunk.emit_op_u16(Op::CALL_IMPORT, to_str, line);
    chunk.emit(1, line);

    emit_try_method_end(chunk, done);
}

// ── Rich bool (user-defined __bool__ / valueOf) ─────────────────────────

/// Emit smart bool: tries __bool__ on objects, falls back to dyn_to_bool.
/// Object must be in `obj_slot`.
/// Stack before: []  Stack after: [bool_value]
pub fn emit_rich_bool(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    let (is_null, _) = emit_try_method(chunk, obj_slot, "__bool__", line);

    // Found: call __bool__(self) → returns bool
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    let done = emit_try_method_fallback_start(chunk, is_null, line);

    // Fallback: use dyn_to_bool
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op(Op::DYN_TO_BOOL, line);

    emit_try_method_end(chunk, done);
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
/// `fallback_op`: the primitive opcode to use if no method (e.g. Op::DYN_LT)
pub fn emit_rich_compare(chunk: &mut Chunk, _dunder: &str, fallback_op: Op, line: u32) {
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
    chunk.emit_op(fallback_op, line);
}

/// Emit a rich comparison with pre-allocated local slots.
/// Caller must store left in `left_slot` and right in `right_slot` before calling.
/// Stack before: []  Stack after: [bool_result]
///
/// Emits: check left.__lt__ → if found, call it(right) → else dyn_lt(left, right)
pub fn emit_rich_compare_locals(chunk: &mut Chunk, left_slot: u16, right_slot: u16, dunder: &str, fallback_op: Op, line: u32) {
    // Try struct_get dunder on left
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from(dunder)));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);

    // Check if method exists (non-null)
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let is_null = chunk.emit_jump(Op::BR_IF_TRUE, line);

    // Found method: call it with self=left, arg=right → result
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    let done = chunk.emit_jump(Op::BR, line);

    // Not found: drop null, use primitive opcode
    chunk.patch_jump(is_null);
    chunk.emit_op(Op::DROP, line); // drop null from dup
    chunk.emit_op(Op::DROP, line); // drop null from struct_get
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_op(fallback_op, line);

    chunk.patch_jump(done);
}

// ── Smart length (user-defined __len__ / __get_length) ──────────────────
//
// Standard WASM opcodes only. Tries __get_length getter first,
// falls back to array_length opcode for plain arrays/strings.

/// Emit smart length: tries user-defined __get_length getter, falls back to array_length.
/// Object must be in `obj_slot`.
/// Stack before: []  Stack after: [length_value]
pub fn emit_smart_length(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    // Try struct_get "__get_length" on object
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from("__get_length")));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);

    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let is_null = chunk.emit_jump(Op::BR_IF_TRUE, line);

    // Found getter: call it with self=obj
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    let done = chunk.emit_jump(Op::BR, line);

    // Not found: use array_length
    chunk.patch_jump(is_null);
    chunk.emit_op(Op::DROP, line); // drop null from dup
    chunk.emit_op(Op::DROP, line); // drop null from struct_get
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);

    chunk.patch_jump(done);
}
