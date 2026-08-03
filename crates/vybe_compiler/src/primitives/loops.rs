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

use crate::primitives::instructions::core_wasm;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

// ── Basic loop primitives ──────────────────────────────────────────────
//
// These are the building blocks for while, do-while, and C-style for loops.
// All compilers MUST use these instead of hand-rolling loop bytecode.

/// Emit the start of a while loop using WASM structured control flow.
/// Emits: block { loop { ... }}
/// Returns (block_patch, loop_patch) — caller compiles condition+body, then calls emit_loop_end.
///
/// Usage:
///   let lp = emit_loop_start(chunks, current, line);
///   // caller: compile condition expression
///   let _ = emit_loop_cond(chunks, current, line);
///   // caller: compile body
///   emit_loop_end(chunks, current, lp, line);
pub struct LoopState {
    pub block_patch: usize,
    pub loop_patch: usize,
    /// If set, there's an inner body block for continue-to-increment pattern.
    /// continue = `br 0` (body block), break = `br 2` (outer block)
    /// If None, continue = `br 0` (loop), break = `br 1` (outer block)
    pub body_block_patch: Option<usize> }

impl LoopState {
    /// Depth for `continue` (restart loop or skip to increment)
    pub fn continue_depth(&self, nesting_offset: u8) -> u8 {
        // body_block is innermost if present
        nesting_offset
    }
    /// Depth for `break` (exit block/loop entirely)
    pub fn break_depth(&self, nesting_offset: u8) -> u8 {
        if self.loop_patch == 0 {
            nesting_offset // block-only: break targets the block itself (depth 0)
        } else {
            let levels = if self.body_block_patch.is_some() {
                3
            } else {
                2
            };
            nesting_offset + levels - 1
        }
    }
    /// Number of label stack entries this loop occupies
    pub fn depth_count(&self) -> u8 {
        if self.loop_patch == 0 {
            1 // block-only (e.g. switch)
        } else if self.body_block_patch.is_some() {
            3 // block + loop + body block
        } else {
            2 // block + loop
        }
    }
}

pub fn emit_loop_start(chunks: &mut [Chunk], current: usize, line: u32) -> LoopState {
    let block_patch = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    LoopState {
        block_patch,
        loop_patch,
        body_block_patch: None }
}

/// After condition is on stack: convert to bool, branch out of block if false.
pub fn emit_loop_cond(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    // br_if_label 1 = break out of loop to block end (depth 0=loop, 1=block)
    chunks[current].emit_br_if(1, line);
}

/// End of loop: branch back to loop start, emit END for loop and block.
pub fn emit_loop_end(chunks: &mut [Chunk], current: usize, state: LoopState, line: u32) {
    // br_label 0 = continue loop (jump to loop start)
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line); // end loop
    chunks[current].patch_loop(state.loop_patch);
    chunks[current].emit_end(line); // end block
    chunks[current].patch_block(state.block_patch);
}

/// Emit start of a do-while loop using structured CF.
/// Returns LoopState — caller compiles body+condition, then calls emit_do_loop_end.
pub fn emit_do_loop_start(chunks: &mut [Chunk], current: usize, line: u32) -> LoopState {
    let block_patch = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    // A third, innermost block wraps the BODY so that `continue` lands on the
    // condition rather than the top of the body. Branching to the loop label
    // would re-enter the body without testing the condition — `do { … continue;
    // … } while (cond)` must still evaluate `cond` on every iteration, in every
    // language that has the form.
    let body_block_patch = chunks[current].emit_block(line);
    LoopState {
        block_patch,
        loop_patch,
        body_block_patch: Some(body_block_patch) }
}

/// Close the do-loop's body block. Call after the body and BEFORE compiling the
/// condition, so `continue` (a branch out of that block) arrives exactly here.
pub fn emit_do_loop_body_end(chunks: &mut [Chunk], current: usize, state: &LoopState, line: u32) {
    if let Some(patch) = state.body_block_patch {
        chunks[current].emit_end(line);
        chunks[current].patch_block(patch);
    }
}

/// End of do-while: condition on stack, branch back to loop if true.
/// `negate` = true for `until` (loop while condition is FALSE).
pub fn emit_do_loop_end(
    chunks: &mut [Chunk],
    current: usize,
    state: LoopState,
    negate: bool,
    line: u32,
) {
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    if negate {
        crate::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    }
    // br_if_label 0 = continue loop if condition is true
    chunks[current].emit_br_if(0, line);
    chunks[current].emit_end(line); // end loop
    chunks[current].patch_loop(state.loop_patch);
    chunks[current].emit_end(line); // end block
    chunks[current].patch_block(state.block_patch);
}

// ── For-in iteration ────────────────────────────────────────────────────

/// Emit the start of a for-in loop: init index, check condition, load element.
/// Caller must have stored the iterable in `arr_slot` before calling this.
/// Returns (loop_start, exit_jump) — caller must pass these to `emit_for_in_end`.
///
/// Stack after: [element] on top (caller assigns to loop variable)
pub fn emit_for_in_start(
    chunks: &mut [Chunk],
    current: usize,
    arr_slot: u16,
    idx_slot: u16,
    line: u32,
) -> LoopState {
    // i = 0
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    // LOCAL_SET leaves the assigned value on the stack in Vybe bytecode;
    // drop it here so every iteration starts with a clean stack.

    // block $exit { loop $loop {
    let block_patch = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);

    // while i < arr.length — use ecma:array.length (handles Array AND
    // TypedArray) rather than the raw Op::ARRAY_LENGTH GC opcode which
    // only handles ObjectKind::Array. This ensures HOF polyfills work
    // uniformly for both plain arrays and typed-array receivers.
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    {
        let idx = chunks[current].add_import("ecma:array", "length");
        chunks[current].emit_call(idx, 1, line);
    }
    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    emit_loop_cond(chunks, current, line);

    // block $body { — continue targets this, skips to increment
    let body_block_patch = chunks[current].emit_block(line);

    // element = arr[i]
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line);

    LoopState {
        block_patch,
        loop_patch,
        body_block_patch: Some(body_block_patch) }
}

/// Emit the end of a for-in loop: increment index, continue loop, close block+loop.
pub fn emit_for_in_end(
    chunks: &mut [Chunk],
    current: usize,
    idx_slot: u16,
    state: LoopState,
    line: u32,
) {
    // Close body block (continue lands here, before increment)
    if let Some(bp) = state.body_block_patch {
        chunks[current].emit_end(line);
        chunks[current].patch_block(bp);
    }

    // i += 1 (increment — runs after continue)
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    // LOCAL_SET peeks (Vybe convention) — drop the residue so the
    // stack height at loop top is invariant across iterations.

    // br $loop (continue loop)
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line); // end loop
    chunks[current].patch_loop(state.loop_patch);
    chunks[current].emit_end(line); // end block
    chunks[current].patch_block(state.block_patch);
}

// ── Map ─────────────────────────────────────────────────────────────────

/// Emit `map(fn, arr)` → new array with fn(x) for each x in arr.
/// Caller must have stored fn in `fn_slot` and array in `arr_slot`.
/// Stack after: [result_array]
pub fn emit_map(
    chunks: &mut [Chunk],
    current: usize,
    fn_slot: u16,
    arr_slot: u16,
    result_slot: u16,
    idx_slot: u16,
    line: u32,
) {
    // result = []
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = emit_for_in_start(chunks, current, arr_slot, idx_slot, line);

    // Drop element from for_in_start — we'll re-fetch inline
    chunks[current].emit_op(Op::DROP, line);

    // result.push(fn(arr[i], i))
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Emit `filter(fn, arr)` → new array with elements where fn(x) is true.
/// Caller must have stored fn in `fn_slot` and array in `arr_slot`.
/// Stack after: [result_array]
pub fn emit_filter(
    chunks: &mut [Chunk],
    current: usize,
    fn_slot: u16,
    arr_slot: u16,
    result_slot: u16,
    idx_slot: u16,
    elem_slot: u16,
    line: u32,
) {
    // result = []
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = emit_for_in_start(chunks, current, arr_slot, idx_slot, line);

    // Store element
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // if fn(element): result.push(element)
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    // Use structured if for the conditional push
    let if_block = chunks[current].emit_block(line);
    crate::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // skip push if false

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line); // end if block
    chunks[current].patch_block(if_block);

    emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Emit `forEach(fn, arr)` → call fn(x) for each x, returns null.
/// Stack after: [null]
pub fn emit_foreach(
    chunks: &mut [Chunk],
    current: usize,
    fn_slot: u16,
    arr_slot: u16,
    idx_slot: u16,
    line: u32,
) {
    let state = emit_for_in_start(chunks, current, arr_slot, idx_slot, line);

    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Emit `reduce(fn, arr)` → fn(fn(arr[0], arr[1]), arr[2]), ...
/// Stack after: [accumulated_value]
pub fn emit_reduce(
    chunks: &mut [Chunk],
    current: usize,
    fn_slot: u16,
    arr_slot: u16,
    acc_slot: u16,
    idx_slot: u16,
    line: u32,
) {
    // acc = arr[0]
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_i32_const(0, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_slot, line);

    // i = 1
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    // block { loop {
    let state = emit_loop_start(chunks, current, line);

    // while i < arr.length — use ecma:array.length (handles TypedArray too).
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    {
        let idx = chunks[current].add_import("ecma:array", "length");
        chunks[current].emit_call(idx, 1, line);
    }
    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    emit_loop_cond(chunks, current, line);

    // acc = fn(acc, arr[i], i)  — ECMA-262 §23.1.3.26: callback(acc, elem, index, array)
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_slot, line);

    // i += 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_slot, line);
}

/// Emit `any(fn, arr)` → true if fn(x) is true for any x.
/// Emit `every(fn, arr)` → true if fn(x) is true for all x.
/// `is_any=true` for any(), `is_any=false` for every().
/// Stack after: [bool]
pub fn emit_any_every(
    chunks: &mut [Chunk],
    current: usize,
    fn_slot: u16,
    arr_slot: u16,
    idx_slot: u16,
    is_any: bool,
    line: u32,
) {
    // Implements: arr.any(fn) / arr.every(fn) as INLINE bytecode that leaves
    // a single bool on the stack — must NOT use Op::return because that
    // aborts the enclosing user function.
    //
    // Pattern:
    //   for elem in arr:
    //     if fn(elem) (any) → push true, jump to end
    //     if !fn(elem) (every) → push false, jump to end
    //   (loop fell through with no match) → push false (any) or true (every)
    let result_local = idx_slot + 1; // assume caller allocated enough locals

    // Set default result BEFORE the loop
    if is_any {
        chunks[current].emit_bool_const(false, line);
    } else {
        chunks[current].emit_bool_const(true, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_local, line);

    let state = emit_for_in_start(chunks, current, arr_slot, idx_slot, line);

    // Drop element from for_in_start, call fn(arr[i]) directly
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    // Structure from emit_for_in_start: block $exit { loop $loop { cond, block $body {
    // From here: depth 0=$body, 1=$loop, 2=$exit
    // With an extra block $skip: depth 0=$skip, 1=$body, 2=$loop, 3=$exit
    if is_any {
        let skip = chunks[current].emit_block(line);
        crate::primitives::ops::emit_dyn_not(&mut chunks[current], line);
        chunks[current].emit_br_if(0, line); // skip if false
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_local, line);
        chunks[current].emit_br(3, line); // break: skip=0, body=1, loop=2, exit=3
        chunks[current].emit_end(line);
        chunks[current].patch_block(skip);
    } else {
        let skip = chunks[current].emit_block(line);
        chunks[current].emit_br_if(0, line); // skip if true
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_local, line);
        chunks[current].emit_br(3, line); // break: skip=0, body=1, loop=2, exit=3
        chunks[current].emit_end(line);
        chunks[current].patch_block(skip);
    }

    emit_for_in_end(chunks, current, idx_slot, state, line);

    // Result already set (default or early exit override) — push it
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_local, line);
}
