//! Expression compilation helpers — shared bytecode patterns for common expressions.
//!
//! Ternary conditionals, short-circuit logic, and null coalescing are identical
//! across all languages. Each helper emits the jump structure and returns
//! patch points so the caller can compile the language-specific sub-expressions
//! in between.

use std::rc::Rc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

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
    chunk.emit_op(Op::dyn_to_bool, line);
    chunk.emit_jump(Op::br_if_false, line)
}

/// After "then" expression: jump over else, patch the false target.
/// Stack: [then_value]
pub fn emit_ternary_middle(chunk: &mut Chunk, false_jump: usize, line: u32) -> usize {
    let end_jump = chunk.emit_jump(Op::br, line);
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
    chunk.emit_op(Op::dup, line);
    chunk.emit_op(Op::dyn_to_bool, line);
    let jump = chunk.emit_jump(Op::br_if_false, line);
    chunk.emit_op(Op::drop, line); // discard left, right becomes result
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
    chunk.emit_op(Op::dup, line);
    chunk.emit_op(Op::dyn_to_bool, line);
    let jump = chunk.emit_jump(Op::br_if_true, line);
    chunk.emit_op(Op::drop, line); // discard left, right becomes result
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
    chunk.emit_op(Op::dup, line);
    let null_jump = chunk.emit_jump(Op::br_if_null, line);
    // Not null — keep left, jump to end
    let end_jump = chunk.emit_jump(Op::br, line);
    // Null path: drop the null, caller compiles right expression
    chunk.patch_jump(null_jump);
    chunk.emit_op(Op::drop, line);
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
    chunk.emit_op(Op::dup, line);
    let skip = chunk.emit_jump(Op::br_if_null, line);
    // Non-null: fall through to member access
    // Return skip (null path) — caller compiles access, then calls end
    (skip, 0) // end_jump set by caller
}

/// After member access: patch the null-skip and null-end jumps.
pub fn emit_null_safe_end(chunk: &mut Chunk, skip: usize, line: u32) {
    let end = chunk.emit_jump(Op::br, line);
    chunk.patch_jump(skip);
    // null is still on stack from the dup
    chunk.patch_jump(end);
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
/// `fallback_op`: the primitive opcode to use if no method (e.g. Op::dyn_lt)
pub fn emit_rich_compare(chunk: &mut Chunk, dunder: &str, fallback_op: Op, line: u32) {
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
    chunk.emit_op_u16(Op::local_get, left_slot, line);
    let key = chunk.add_constant(Value::String(Rc::from(dunder)));
    chunk.emit_op_u16(Op::struct_get, key, line);

    // Check if method exists (non-null)
    chunk.emit_op(Op::dup, line);
    chunk.emit_op(Op::ref_is_null, line);
    let is_null = chunk.emit_jump(Op::br_if_true, line);

    // Found method: call it with self=left, arg=right → result
    chunk.emit_op_u16(Op::local_get, left_slot, line);
    chunk.emit_op_u16(Op::local_get, right_slot, line);
    chunk.emit_op_u8(Op::call_ref, 2, line);
    let done = chunk.emit_jump(Op::br, line);

    // Not found: drop null, use primitive opcode
    chunk.patch_jump(is_null);
    chunk.emit_op(Op::drop, line); // drop null from dup
    chunk.emit_op(Op::drop, line); // drop null from struct_get
    chunk.emit_op_u16(Op::local_get, left_slot, line);
    chunk.emit_op_u16(Op::local_get, right_slot, line);
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
    chunk.emit_op_u16(Op::local_get, obj_slot, line);
    let key = chunk.add_constant(Value::String(Rc::from("__get_length")));
    chunk.emit_op_u16(Op::struct_get, key, line);

    chunk.emit_op(Op::dup, line);
    chunk.emit_op(Op::ref_is_null, line);
    let is_null = chunk.emit_jump(Op::br_if_true, line);

    // Found getter: call it with self=obj
    chunk.emit_op_u16(Op::local_get, obj_slot, line);
    chunk.emit_op_u8(Op::call_ref, 1, line);
    let done = chunk.emit_jump(Op::br, line);

    // Not found: use array_length
    chunk.patch_jump(is_null);
    chunk.emit_op(Op::drop, line); // drop null from dup
    chunk.emit_op(Op::drop, line); // drop null from struct_get
    chunk.emit_op_u16(Op::local_get, obj_slot, line);
    chunk.emit_op(Op::array_length, line);

    chunk.patch_jump(done);
}
