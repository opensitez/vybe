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

/// Emit the start of a while loop using WASM structured control flow.
/// Emits: block { loop { ... }}
/// Returns (block_patch, loop_patch) — caller compiles condition+body, then calls emit_loop_end.
///
/// Usage:
///   let lp = emit_loop_start(chunk, line);
///   // caller: compile condition expression
///   let _ = emit_loop_cond(chunk, line);
///   // caller: compile body
///   emit_loop_end(chunk, lp, line);
pub struct LoopState {
    pub block_patch: usize,
    pub loop_patch: usize,
    /// If set, there's an inner body block for continue-to-increment pattern.
    /// continue = BR_LABEL 0 (body block), break = BR_LABEL 2 (outer block)
    /// If None, continue = BR_LABEL 0 (loop), break = BR_LABEL 1 (outer block)
    pub body_block_patch: Option<usize>,
}

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
            let levels = if self.body_block_patch.is_some() { 3 } else { 2 };
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

pub fn emit_loop_start(chunk: &mut Chunk, line: u32) -> LoopState {
    let block_patch = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    LoopState { block_patch, loop_patch, body_block_patch: None }
}

/// After condition is on stack: convert to bool, branch out of block if false.
pub fn emit_loop_cond(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    chunk.emit_op(Op::DYN_NOT, line);
    // br_if_label 1 = break out of loop to block end (depth 0=loop, 1=block)
    chunk.emit_br_if(1, line);
}

/// End of loop: branch back to loop start, emit END for loop and block.
pub fn emit_loop_end(chunk: &mut Chunk, state: LoopState, line: u32) {
    // br_label 0 = continue loop (jump to loop start)
    chunk.emit_br(0, line);
    chunk.emit_end(line); // end loop
    chunk.patch_loop(state.loop_patch);
    chunk.emit_end(line); // end block
    chunk.patch_block(state.block_patch);
}

/// Emit start of a do-while loop using structured CF.
/// Returns LoopState — caller compiles body+condition, then calls emit_do_loop_end.
pub fn emit_do_loop_start(chunk: &mut Chunk, line: u32) -> LoopState {
    let block_patch = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    LoopState { block_patch, loop_patch, body_block_patch: None }
}

/// End of do-while: condition on stack, branch back to loop if true.
/// `negate` = true for `until` (loop while condition is FALSE).
pub fn emit_do_loop_end(chunk: &mut Chunk, state: LoopState, negate: bool, line: u32) {
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    if negate {
        chunk.emit_op(Op::DYN_NOT, line);
    }
    // br_if_label 0 = continue loop if condition is true
    chunk.emit_br_if(0, line);
    chunk.emit_end(line); // end loop
    chunk.patch_loop(state.loop_patch);
    chunk.emit_end(line); // end block
    chunk.patch_block(state.block_patch);
}

// ── For-in iteration ────────────────────────────────────────────────────

/// Emit the start of a for-in loop: init index, check condition, load element.
/// Caller must have stored the iterable in `arr_slot` before calling this.
/// Returns (loop_start, exit_jump) — caller must pass these to `emit_for_in_end`.
///
/// Stack after: [element] on top (caller assigns to loop variable)
pub fn emit_for_in_start(chunk: &mut Chunk, arr_slot: u16, idx_slot: u16, line: u32) -> LoopState {
    // i = 0
    chunk.emit_op(Op::I32_CONST_0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    // block $exit { loop $loop {
    let block_patch = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);

    // while i < arr.length
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op(Op::DYN_LT, line);
    emit_loop_cond(chunk, line);

    // block $body { — continue targets this, skips to increment
    let body_block_patch = chunk.emit_block(line);

    // element = arr[i]
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);

    LoopState { block_patch, loop_patch, body_block_patch: Some(body_block_patch) }
}

/// Emit the end of a for-in loop: increment index, continue loop, close block+loop.
pub fn emit_for_in_end(chunk: &mut Chunk, idx_slot: u16, state: LoopState, line: u32) {
    // Close body block (continue lands here, before increment)
    if let Some(bp) = state.body_block_patch {
        chunk.emit_end(line);
        chunk.patch_block(bp);
    }

    // i += 1 (increment — runs after continue)
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::I32_CONST_1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    // br $loop (continue loop)
    chunk.emit_br(0, line);
    chunk.emit_end(line); // end loop
    chunk.patch_loop(state.loop_patch);
    chunk.emit_end(line); // end block
    chunk.patch_block(state.block_patch);
}

// ── Map ─────────────────────────────────────────────────────────────────

/// Emit `map(fn, arr)` → new array with fn(x) for each x in arr.
/// Caller must have stored fn in `fn_slot` and array in `arr_slot`.
/// Stack after: [result_array]
pub fn emit_map(chunk: &mut Chunk, fn_slot: u16, arr_slot: u16, result_slot: u16, idx_slot: u16, line: u32) {
    // result = []
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_op(Op::DROP, line);

    let state = emit_for_in_start(chunk, arr_slot, idx_slot, line);

    // Drop element from for_in_start — we'll re-fetch inline
    chunk.emit_op(Op::DROP, line);

    // result.push(fn(arr[i], i))
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_op(Op::ARRAY_PUSH, line);
    chunk.emit_op(Op::DROP, line);

    emit_for_in_end(chunk, idx_slot, state, line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Emit `filter(fn, arr)` → new array with elements where fn(x) is true.
/// Caller must have stored fn in `fn_slot` and array in `arr_slot`.
/// Stack after: [result_array]
pub fn emit_filter(chunk: &mut Chunk, fn_slot: u16, arr_slot: u16, result_slot: u16, idx_slot: u16, elem_slot: u16, line: u32) {
    // result = []
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_op(Op::DROP, line);

    let state = emit_for_in_start(chunk, arr_slot, idx_slot, line);

    // Store element
    chunk.emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunk.emit_op(Op::DROP, line);

    // if fn(element): result.push(element)
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    // Use structured if for the conditional push
    let if_block = chunk.emit_block(line);
    chunk.emit_op(Op::DYN_NOT, line);
    chunk.emit_br_if(0, line); // skip push if false

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunk.emit_op(Op::ARRAY_PUSH, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line); // end if block
    chunk.patch_block(if_block);

    emit_for_in_end(chunk, idx_slot, state, line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Emit `forEach(fn, arr)` → call fn(x) for each x, returns null.
/// Stack after: [null]
pub fn emit_foreach(chunk: &mut Chunk, fn_slot: u16, arr_slot: u16, idx_slot: u16, line: u32) {
    let state = emit_for_in_start(chunk, arr_slot, idx_slot, line);

    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_op(Op::DROP, line);

    emit_for_in_end(chunk, idx_slot, state, line);

    chunk.emit_op(Op::NULL, line);
}

/// Emit `reduce(fn, arr)` → fn(fn(arr[0], arr[1]), arr[2]), ...
/// Stack after: [accumulated_value]
pub fn emit_reduce(chunk: &mut Chunk, fn_slot: u16, arr_slot: u16, acc_slot: u16, idx_slot: u16, line: u32) {
    
    use vybe_bytecode::Value;

    // acc = arr[0]
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    let zero = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, zero, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc_slot, line);
    chunk.emit_op(Op::DROP, line);

    // i = 1
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, one, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    // block { loop {
    let state = emit_loop_start(chunk, line);

    // while i < arr.length
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op(Op::DYN_LT, line);
    emit_loop_cond(chunk, line);

    // acc = fn(acc, arr[i])
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, acc_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc_slot, line);
    chunk.emit_op(Op::DROP, line);

    // i += 1
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::I32_CONST_1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    emit_loop_end(chunk, state, line);

    chunk.emit_op_u16(Op::LOCAL_GET, acc_slot, line);
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
    let result_local = idx_slot + 1; // assume caller allocated enough locals

    // Set default result BEFORE the loop
    if is_any { chunk.emit_op(Op::FALSE, line); } else { chunk.emit_op(Op::TRUE, line); }
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
    chunk.emit_op(Op::DROP, line);

    let state = emit_for_in_start(chunk, arr_slot, idx_slot, line);

    // Drop element from for_in_start, call fn(arr[i]) directly
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    // Structure from emit_for_in_start: block $exit { loop $loop { cond, block $body {
    // From here: depth 0=$body, 1=$loop, 2=$exit
    // With an extra block $skip: depth 0=$skip, 1=$body, 2=$loop, 3=$exit
    if is_any {
        let skip = chunk.emit_block(line);
        chunk.emit_op(Op::DYN_NOT, line);
        chunk.emit_br_if(0, line); // skip if false
        chunk.emit_op(Op::TRUE, line);
        chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_br(3, line); // break: skip=0, body=1, loop=2, exit=3
        chunk.emit_end(line);
        chunk.patch_block(skip);
    } else {
        let skip = chunk.emit_block(line);
        chunk.emit_br_if(0, line); // skip if true
        chunk.emit_op(Op::FALSE, line);
        chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_br(3, line); // break: skip=0, body=1, loop=2, exit=3
        chunk.emit_end(line);
        chunk.patch_block(skip);
    }

    emit_for_in_end(chunk, idx_slot, state, line);

    // Result already set (default or early exit override) — push it
    chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
}
