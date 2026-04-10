//! Loop and iteration helpers — shared bytecode patterns for collection operations.
//!
//! All Vybe compilers emit identical bytecode for map/filter/forEach/reduce/any/every
//! and for-in iteration. This module centralizes those patterns so every language
//! produces compatible bytecode.
//!
//! ## Slot contract
//!
//! Each function takes pre-allocated local slots. The caller is responsible for:
//! 1. Allocating slots via their scope system
//! 2. Storing the function/array values into those slots BEFORE calling these helpers
//! 3. The helpers only emit the loop body — not the argument evaluation

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

// ── Basic loop primitives ──────────────────────────────────────────────
//
// These are the building blocks for while, do-while, and C-style for loops.
// All compilers MUST use these instead of hand-rolling loop bytecode.

/// Emit the start of a while loop: mark loop start, compile condition already
/// on stack, branch out if false.
/// Returns (loop_start, exit_jump) — caller compiles body, then calls emit_loop_end.
///
/// Usage:
///   let (start, exit) = emit_while_start(chunk, line);
///   // caller: compile condition expression
///   // caller: emit dyn_to_bool
///   // caller: let exit = emit_while_cond(chunk, line);
///   // caller: compile body
///   emit_loop_end(chunk, start, exit, line);
pub fn emit_loop_start(chunk: &mut Chunk) -> usize {
    chunk.current_offset()
}

/// After condition is on stack: convert to bool, jump out if false.
/// Returns exit_jump to patch later.
pub fn emit_loop_cond(chunk: &mut Chunk, line: u32) -> usize {
    chunk.emit_op(Op::dyn_to_bool, line);
    chunk.emit_jump(Op::br_if_false, line)
}

/// End of loop: jump back to start, patch the exit.
pub fn emit_loop_end(chunk: &mut Chunk, loop_start: usize, exit_jump: usize, line: u32) {
    chunk.emit_loop(loop_start, line);
    chunk.patch_jump(exit_jump);
}

/// Emit unconditional loop back (for do-while where condition is at the end).
/// Returns loop_start for the unconditional loop point.
pub fn emit_do_loop_start(chunk: &mut Chunk) -> usize {
    chunk.current_offset()
}

/// End of do-while: condition on stack, branch back to start if true.
/// `negate` = true for `until` (loop while condition is FALSE).
pub fn emit_do_loop_end(chunk: &mut Chunk, loop_start: usize, negate: bool, line: u32) {
    chunk.emit_op(Op::dyn_to_bool, line);
    if negate {
        chunk.emit_op(Op::dyn_not, line);
    }
    // Branch back to start if condition is true
    let exit = chunk.emit_jump(Op::br_if_false, line);
    chunk.emit_loop(loop_start, line);
    chunk.patch_jump(exit);
}

// ── For-in iteration ────────────────────────────────────────────────────

/// Emit the start of a for-in loop: init index, check condition, load element.
/// Caller must have stored the iterable in `arr_slot` before calling this.
/// Returns (loop_start, exit_jump) — caller must pass these to `emit_for_in_end`.
///
/// Stack after: [element] on top (caller assigns to loop variable)
pub fn emit_for_in_start(chunk: &mut Chunk, arr_slot: u16, idx_slot: u16, line: u32) -> (usize, usize) {
    // i = 0
    chunk.emit_op(Op::i32_const_0, line);
    chunk.emit_op_u16(Op::local_set, idx_slot, line);

    let loop_start = chunk.current_offset();

    // while i < arr.length
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op_u16(Op::local_get, arr_slot, line);
    chunk.emit_op(Op::array_length, line);
    chunk.emit_op(Op::dyn_lt, line);
    let exit_jump = chunk.emit_jump(Op::br_if_false, line);

    // element = arr[i]
    chunk.emit_op_u16(Op::local_get, arr_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line);

    (loop_start, exit_jump)
}

/// Emit the end of a for-in loop: increment index, loop back, patch exit.
pub fn emit_for_in_end(chunk: &mut Chunk, idx_slot: u16, loop_start: usize, exit_jump: usize, line: u32) {
    // i += 1
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::i32_const_1, line);
    chunk.emit_op(Op::i32_add, line);
    chunk.emit_op_u16(Op::local_set, idx_slot, line);

    chunk.emit_loop(loop_start, line);
    chunk.patch_jump(exit_jump);
}

// ── Map ─────────────────────────────────────────────────────────────────

/// Emit `map(fn, arr)` → new array with fn(x) for each x in arr.
/// Caller must have stored fn in `fn_slot` and array in `arr_slot`.
/// Stack after: [result_array]
pub fn emit_map(chunk: &mut Chunk, fn_slot: u16, arr_slot: u16, result_slot: u16, idx_slot: u16, line: u32) {
    // result = []
    chunk.emit_op_u16(Op::array_new, 0, line);
    chunk.emit_op_u16(Op::local_set, result_slot, line);
    chunk.emit_op(Op::drop, line);

    let (loop_start, exit_jump) = emit_for_in_start(chunk, arr_slot, idx_slot, line);

    // Stack: [element]. Call fn(element), push result.
    let elem_slot = idx_slot; // reuse naming — element is on stack, not in slot
    // Actually we need: result.push(fn(arr[i]))
    // Stack has [element] from for_in_start. Store it, then build call.
    // Simpler: rewrite to not use for_in_start since we need the fn call pattern.

    // Drop element from for_in_start — we'll re-fetch inline
    chunk.emit_op(Op::drop, line);

    // result.push(fn(arr[i]))
    chunk.emit_op_u16(Op::local_get, result_slot, line);
    chunk.emit_op_u16(Op::local_get, fn_slot, line);
    chunk.emit_op_u16(Op::local_get, arr_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line);
    chunk.emit_op_u8(Op::call_ref, 1, line);
    chunk.emit_op(Op::array_push, line);
    chunk.emit_op(Op::drop, line);

    emit_for_in_end(chunk, idx_slot, loop_start, exit_jump, line);

    chunk.emit_op_u16(Op::local_get, result_slot, line);
}

/// Emit `filter(fn, arr)` → new array with elements where fn(x) is true.
/// Caller must have stored fn in `fn_slot` and array in `arr_slot`.
/// Stack after: [result_array]
pub fn emit_filter(chunk: &mut Chunk, fn_slot: u16, arr_slot: u16, result_slot: u16, idx_slot: u16, elem_slot: u16, line: u32) {
    // result = []
    chunk.emit_op_u16(Op::array_new, 0, line);
    chunk.emit_op_u16(Op::local_set, result_slot, line);
    chunk.emit_op(Op::drop, line);

    let (loop_start, exit_jump) = emit_for_in_start(chunk, arr_slot, idx_slot, line);

    // Store element
    chunk.emit_op_u16(Op::local_set, elem_slot, line);
    chunk.emit_op(Op::drop, line);

    // if fn(element): result.push(element)
    chunk.emit_op_u16(Op::local_get, fn_slot, line);
    chunk.emit_op_u16(Op::local_get, elem_slot, line);
    chunk.emit_op_u8(Op::call_ref, 1, line);
    chunk.emit_op(Op::dyn_to_bool, line);
    let skip_push = chunk.emit_jump(Op::br_if_false, line);

    chunk.emit_op_u16(Op::local_get, result_slot, line);
    chunk.emit_op_u16(Op::local_get, elem_slot, line);
    chunk.emit_op(Op::array_push, line);
    chunk.emit_op(Op::drop, line);

    chunk.patch_jump(skip_push);

    emit_for_in_end(chunk, idx_slot, loop_start, exit_jump, line);

    chunk.emit_op_u16(Op::local_get, result_slot, line);
}

/// Emit `forEach(fn, arr)` → call fn(x) for each x, returns null.
/// Stack after: [null]
pub fn emit_foreach(chunk: &mut Chunk, fn_slot: u16, arr_slot: u16, idx_slot: u16, line: u32) {
    let (loop_start, exit_jump) = emit_for_in_start(chunk, arr_slot, idx_slot, line);

    // Drop element from for_in_start, call fn(arr[i]) directly
    chunk.emit_op(Op::drop, line);

    chunk.emit_op_u16(Op::local_get, fn_slot, line);
    chunk.emit_op_u16(Op::local_get, arr_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line);
    chunk.emit_op_u8(Op::call_ref, 1, line);
    chunk.emit_op(Op::drop, line);

    emit_for_in_end(chunk, idx_slot, loop_start, exit_jump, line);

    chunk.emit_op(Op::null, line);
}

/// Emit `reduce(fn, arr)` → fn(fn(arr[0], arr[1]), arr[2]), ...
/// Stack after: [accumulated_value]
pub fn emit_reduce(chunk: &mut Chunk, fn_slot: u16, arr_slot: u16, acc_slot: u16, idx_slot: u16, line: u32) {
    use std::sync::Arc;
    use vybe_bytecode::Value;

    // acc = arr[0]
    chunk.emit_op_u16(Op::local_get, arr_slot, line);
    let zero = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::r#const, zero, line);
    chunk.emit_op(Op::array_get, line);
    chunk.emit_op_u16(Op::local_set, acc_slot, line);
    chunk.emit_op(Op::drop, line);

    // i = 1
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, one, line);
    chunk.emit_op_u16(Op::local_set, idx_slot, line);

    let loop_start = chunk.current_offset();

    // while i < arr.length
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op_u16(Op::local_get, arr_slot, line);
    chunk.emit_op(Op::array_length, line);
    chunk.emit_op(Op::dyn_lt, line);
    let exit_jump = chunk.emit_jump(Op::br_if_false, line);

    // acc = fn(acc, arr[i])
    chunk.emit_op_u16(Op::local_get, fn_slot, line);
    chunk.emit_op_u16(Op::local_get, acc_slot, line);
    chunk.emit_op_u16(Op::local_get, arr_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line);
    chunk.emit_op_u8(Op::call_ref, 2, line);
    chunk.emit_op_u16(Op::local_set, acc_slot, line);
    chunk.emit_op(Op::drop, line);

    emit_for_in_end(chunk, idx_slot, loop_start, exit_jump, line);

    chunk.emit_op_u16(Op::local_get, acc_slot, line);
}

/// Emit `any(fn, arr)` → true if fn(x) is true for any x.
/// Emit `every(fn, arr)` → true if fn(x) is true for all x.
/// `is_any=true` for any(), `is_any=false` for every().
/// Stack after: [bool]
pub fn emit_any_every(chunk: &mut Chunk, fn_slot: u16, arr_slot: u16, idx_slot: u16, is_any: bool, line: u32) {
    // Implements: arr.any(fn) / arr.every(fn) as INLINE bytecode that leaves
    // a single bool on the stack — must NOT use Op::return because that
    // aborts the enclosing user function.
    //
    // Pattern:
    //   for elem in arr:
    //     if fn(elem) (any) → push true, jump to end
    //     if !fn(elem) (every) → push false, jump to end
    //   (loop fell through with no match) → push false (any) or true (every)
    let (loop_start, exit_jump) = emit_for_in_start(chunk, arr_slot, idx_slot, line);

    // Drop element from for_in_start, call fn(arr[i]) directly
    chunk.emit_op(Op::drop, line);

    chunk.emit_op_u16(Op::local_get, fn_slot, line);
    chunk.emit_op_u16(Op::local_get, arr_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line);
    chunk.emit_op_u8(Op::call_ref, 1, line);
    chunk.emit_op(Op::dyn_to_bool, line);

    // Patches that jump to the "found a match — leave result on stack" arm.
    let mut early_exit_patches: Vec<usize> = Vec::new();
    if is_any {
        // any: if true → break out with `true`
        let no_match = chunk.emit_jump(Op::br_if_false, line);
        chunk.emit_op(Op::r#true, line);
        early_exit_patches.push(chunk.emit_jump(Op::br, line));
        chunk.patch_jump(no_match);
    } else {
        // every: if false → break out with `false`
        let still_ok = chunk.emit_jump(Op::br_if_true, line);
        chunk.emit_op(Op::r#false, line);
        early_exit_patches.push(chunk.emit_jump(Op::br, line));
        chunk.patch_jump(still_ok);
    }

    emit_for_in_end(chunk, idx_slot, loop_start, exit_jump, line);

    // Loop completed with no early exit → any=false, every=true
    if is_any { chunk.emit_op(Op::r#false, line); } else { chunk.emit_op(Op::r#true, line); }

    // Land here from the early-exit branch with the result already on the stack.
    for p in early_exit_patches { chunk.patch_jump(p); }
}
