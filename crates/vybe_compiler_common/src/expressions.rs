//! Expression compilation helpers — shared bytecode patterns for common expressions.
//!
//! Ternary conditionals, short-circuit logic, and null coalescing are identical
//! across all languages. Each helper emits the jump structure and returns
//! patch points so the caller can compile the language-specific sub-expressions
//! in between.

use vybe_bytecode::Chunk;
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
