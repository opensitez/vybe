//! Collection operations — arrays, sets, sorting, range.
//!
//! Array ops use direct WASM GC opcodes where possible.
//! Higher-level ops (range, sorted, enumerate, zip) use Vybe host imports
//! when available, or fall back to inline WASM bytecode sequences.

use vybe_bytecode::Chunk;
#[allow(unused_imports)]
use vybe_bytecode::Value;
use vybe_bytecode::opcode::Op;
#[allow(unused_imports)]
use crate::Target;

// ── Direct WASM GC opcodes (no host call) ───────────────────

/// Create array from N stack values. Stack: [v1, v2, ..., vN] → [array]
pub fn emit_array_new(chunk: &mut Chunk, count: u16, line: u32) {
    chunk.emit_op_u16(Op::array_new, count, line);
}

/// Array length. Stack: [array] → [i32]
pub fn emit_len(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_length, line);
}

/// Array push. Stack: [array, value] → [array]
pub fn emit_push(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_push, line);
}

/// Array pop. Stack: [array] → [value]
pub fn emit_pop(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_pop, line);
}

/// Array get. Stack: [array, index] → [value]
pub fn emit_get(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_get, line);
}

/// Array set. Stack: [array, index, value] → [value]
pub fn emit_set(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_set, line);
}

/// Array slice. Stack: [array, start, end] → [array]
pub fn emit_slice(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_slice, line);
}

/// Array join. Stack: [array, delimiter] → [string]
pub fn emit_join(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_join, line);
}

/// Array reverse (in-place). Stack: [array] → [array]
pub fn emit_reverse(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_reverse, line);
}

/// Array contains. Stack: [array, value] → [bool]
pub fn emit_contains(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_contains, line);
}

/// Array indexOf. Stack: [array, value] → [i32]
pub fn emit_index_of(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_index_of, line);
}

/// Array concat. Stack: [array, array] → [array]
pub fn emit_concat(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_concat, line);
}

/// Array shift (remove first). Stack: [array] → [value]
pub fn emit_shift(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_shift, line);
}

// ── Host imports (higher-level operations) ──────────────────

/// range(stop) or range(start, stop) or range(start, stop, step).
/// Stack: [args...] → [array]
///
/// On Vybe: single host call. On standard WASM: inline loop.
pub fn emit_range(chunk: &mut Chunk, arg_count: u8, line: u32) {
    // Always use host call — range() requires dynamic allocation that
    // can't be done in pure opcodes without a complex inline loop.
    // On non-Vybe runtimes, this import must be provided by the embedder.
    let idx = chunk.add_import("vybe:array", "range");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(arg_count, line);
}

/// Target-aware range — uses host call on Vybe, inline loop on pure WASM.
/// Stack: [start, stop] → [array]
pub fn emit_range_targeted(chunk: &mut Chunk, arg_count: u8, target: &Target, line: u32) {
    if target.has_module("vybe:array") {
        let idx = chunk.add_import("vybe:array", "range");
        chunk.emit_op_u16(Op::call_import, idx, line);
        chunk.emit(arg_count, line);
    } else {
        // Pure WASM fallback: build array with inline loop.
        // For range(stop): start=0, step=1
        // Stack has [stop] (1 arg) or [start, stop] (2 args)
        // This is complex inline bytecode but portable.
        //
        // Pseudocode:
        //   result = []
        //   i = start
        //   while i < stop: result.push(i); i += 1
        //   return result

        // For simplicity in fallback, we still need array_new + array_push opcodes
        // which ARE standard WASM GC. The loop uses only core WASM control flow.
        if arg_count == 1 {
            // Stack: [stop] — need start=0
            let stop_local = chunk.local_count;
            chunk.local_count += 3; // stop, i, result
            let i_local = stop_local + 1;
            let result_local = stop_local + 2;

            chunk.emit_op_u16(Op::local_set, stop_local, line);  // store stop
            chunk.emit_op(Op::drop, line);
            chunk.emit_op_u16(Op::array_new, 0, line);           // result = []
            chunk.emit_op_u16(Op::local_set, result_local, line);
            chunk.emit_op(Op::drop, line);
            chunk.emit_op(Op::i32_const_0, line);                // i = 0
            chunk.emit_op_u16(Op::local_set, i_local, line);
            chunk.emit_op(Op::drop, line);

            let loop_start = chunk.current_offset();
            chunk.emit_op_u16(Op::local_get, i_local, line);
            chunk.emit_op_u16(Op::local_get, stop_local, line);
            chunk.emit_op(Op::dyn_lt, line);
            let exit = chunk.emit_jump(Op::br_if_false, line);

            chunk.emit_op_u16(Op::local_get, result_local, line);
            chunk.emit_op_u16(Op::local_get, i_local, line);
            chunk.emit_op(Op::array_push, line);
            chunk.emit_op(Op::drop, line);

            chunk.emit_op_u16(Op::local_get, i_local, line);
            chunk.emit_op(Op::i32_const_1, line);
            chunk.emit_op(Op::i32_add, line);
            chunk.emit_op_u16(Op::local_set, i_local, line);
            chunk.emit_op(Op::drop, line);

            chunk.emit_loop(loop_start, line);
            chunk.patch_jump(exit);

            chunk.emit_op_u16(Op::local_get, result_local, line);
        } else {
            // 2+ args: just use host call (complex step handling not worth inlining)
            let idx = chunk.add_import("vybe:array", "range");
            chunk.emit_op_u16(Op::call_import, idx, line);
            chunk.emit(arg_count, line);
        }
    }
}

/// sorted(iterable). Stack: [array] → [sorted_array]
pub fn emit_sorted(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "sorted");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// Target-aware sorted — Vybe host call or standard "env" import.
pub fn emit_sorted_targeted(chunk: &mut Chunk, target: &Target, line: u32) {
    if target.has_module("vybe:array") {
        emit_sorted(chunk, line);
    } else {
        // Standard WASM fallback: require embedder to provide env/sorted.
        // Inlining a sort algorithm as bytecode would bloat every binary.
        // This is the same approach as Emscripten (imports libc functions).
        let idx = chunk.add_import("env", "sorted");
        chunk.emit_op_u16(Op::call_import, idx, line);
        chunk.emit(1, line);
    }
}

/// reversed(iterable). Stack: [array] → [reversed_array]
pub fn emit_reversed(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "reversed");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// enumerate(iterable). Stack: [array] → [array of [i, val] pairs]
pub fn emit_enumerate(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "enumerate");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// zip(a, b). Stack: [array, array] → [array of [a_i, b_i] pairs]
pub fn emit_zip(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "zip");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(2, line);
}

/// sum(iterable). Stack: [array] → [number]
pub fn emit_sum(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "sum");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// min(args...). Stack: [args...] → [min_value]
pub fn emit_min(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("vybe:array", "pymin");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(arg_count, line);
}

/// max(args...). Stack: [args...] → [max_value]
pub fn emit_max(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("vybe:array", "pymax");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(arg_count, line);
}
