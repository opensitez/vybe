//! Collection operations — arrays, sets, sorting, range.
//!
//! Array ops use direct WASM GC opcodes where possible.
//! Higher-level ops (range, sorted, enumerate, zip) use Vybe host imports
//! when available, or fall back to inline WASM bytecode sequences.

use std::sync::Arc;
use vybe_bytecode::Chunk;
use vybe_bytecode::Value;
use vybe_bytecode::opcode::Op;
#[allow(unused_imports)]
use crate::Target;

// ── Direct WASM GC opcodes (no host call) ───────────────────

/// Create array from N stack values. Stack: [v1, v2, ..., vN] → [array]
pub fn emit_array_new(chunk: &mut Chunk, count: u16, line: u32) {
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, count, line);
}

/// Array length. Stack: [array] → [i32]
pub fn emit_len(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_LENGTH, line);
}

/// Array push. Stack: [array, value] → [array]
pub fn emit_push(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_PUSH, line);
}

/// Array pop. Stack: [array] → [value]
pub fn emit_pop(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_POP, line);
}

/// Array get. Stack: [array, index] → [value]
pub fn emit_get(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// Array set. Stack: [array, index, value] → [value]
pub fn emit_set(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_SET, line);
}

/// Raw array slice opcode. Stack: [array, start, end] → [array]
/// Only handles arrays. For polymorphic slicing (string OR array), prefer
/// `emit_slice_push_func` + args + `emit_slice_invoke`, which routes through
/// the stdlib `__vybe_slice` chunk that runtime-dispatches on the value type.
pub fn emit_slice(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_SLICE, line);
}

/// Push the __vybe_slice func ref. Use BEFORE compiling the object/start/end.
/// Pure WASM — bundle wires `__vybe_slice` to `build_slice` stdlib chunk,
/// which dispatches at runtime to either `str_substring` or `array_slice`
/// depending on the operand type. Works uniformly across every language whose
/// surface syntax is `obj[start..end]`.
pub fn emit_slice_push_func(chunk: &mut Chunk, line: u32) {
    let name = chunk.add_constant(Value::String(Arc::from("__vybe_slice")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, line);
}

/// Invoke __vybe_slice after [func, obj, start, end] are on the stack.
pub fn emit_slice_invoke(chunk: &mut Chunk, line: u32) {
    chunk.emit_op_u8(Op::CALL_REF, 3, line);
}

/// Array join. Stack: [array, delimiter] → [string]
pub fn emit_join(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_JOIN, line);
}

/// Array reverse (in-place). Stack: [array] → [array]
pub fn emit_reverse(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_REVERSE, line);
}

/// Array contains. Stack: [array, value] → [bool]
pub fn emit_contains(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_CONTAINS, line);
}

/// Array indexOf. Stack: [array, value] → [i32]
pub fn emit_index_of(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_INDEX_OF, line);
}

/// Array concat. Stack: [array, array] → [array]
pub fn emit_concat(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_CONCAT, line);
}

/// Array shift (remove first). Stack: [array] → [value]
pub fn emit_shift(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ARRAY_SHIFT, line);
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
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(arg_count, line);
}

/// Target-aware range — uses host call on Vybe, inline loop on pure WASM.
/// Stack: [start, stop] → [array]
pub fn emit_range_targeted(chunk: &mut Chunk, arg_count: u8, target: &Target, line: u32) {
    if target.has_module("vybe:array") {
        let idx = chunk.add_import("vybe:array", "range");
        chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
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

            chunk.emit_op_u16(Op::LOCAL_SET, stop_local, line);  // store stop
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);           // result = []
            chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op(Op::I32_CONST_0, line);                // i = 0
            chunk.emit_op_u16(Op::LOCAL_SET, i_local, line);
            chunk.emit_op(Op::DROP, line);

            let loop_start = chunk.current_offset();
            chunk.emit_op_u16(Op::LOCAL_GET, i_local, line);
            chunk.emit_op_u16(Op::LOCAL_GET, stop_local, line);
            chunk.emit_op(Op::DYN_LT, line);
            let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

            chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
            chunk.emit_op_u16(Op::LOCAL_GET, i_local, line);
            chunk.emit_op(Op::ARRAY_PUSH, line);
            chunk.emit_op(Op::DROP, line);

            chunk.emit_op_u16(Op::LOCAL_GET, i_local, line);
            chunk.emit_op(Op::I32_CONST_1, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, i_local, line);
            chunk.emit_op(Op::DROP, line);

            chunk.emit_loop(loop_start, line);
            chunk.patch_jump(exit);

            chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
        } else {
            // 2+ args: just use host call (complex step handling not worth inlining)
            let idx = chunk.add_import("vybe:array", "range");
            chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
            chunk.emit(arg_count, line);
        }
    }
}

/// sorted(iterable). Stack: [array] → [sorted_array]
/// Legacy entry point — uses host import. The bundle aliases vybe:array sorted to __vybe_sorted.
/// Prefer using `emit_sorted_push_func` + args + `emit_sorted_invoke` for pure WASM bytecode.
pub fn emit_sorted(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "sorted");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// Push the __vybe_sorted func ref. Use BEFORE compiling args.
/// Pure WASM — no host import. Bundle wires __vybe_sorted to stdlib chunk.
pub fn emit_sorted_push_func(chunk: &mut Chunk, line: u32) {
    let name = chunk.add_constant(Value::String(Arc::from("__vybe_sorted")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, line);
}

/// Invoke __vybe_sorted after func ref + array are on stack.
pub fn emit_sorted_invoke(chunk: &mut Chunk, line: u32) {
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
}

/// Push the __vybe_sort_in_place func ref. Use BEFORE compiling the array.
/// Pure WASM — no host import. Bundle wires __vybe_sort_in_place to the
/// `build_sort_in_place` stdlib chunk. At runtime, Vybe VM may overwrite
/// the global with an optimized native sort (polyfill pattern).
pub fn emit_sort_in_place_push_func(chunk: &mut Chunk, line: u32) {
    let name = chunk.add_constant(Value::String(Arc::from("__vybe_sort_in_place")));
    chunk.emit_op_u16(Op::GLOBAL_GET, name, line);
}

/// Invoke __vybe_sort_in_place after func ref + array are on stack.
/// Stack: [func, array] → [array] (same reference, mutated in place).
pub fn emit_sort_in_place_invoke(chunk: &mut Chunk, line: u32) {
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
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
        chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
        chunk.emit(1, line);
    }
}

/// reversed(iterable). Stack: [array] → [reversed_array]
pub fn emit_reversed(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "reversed");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// enumerate(iterable). Stack: [array] → [array of [i, val] pairs]
pub fn emit_enumerate(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "enumerate");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// zip(a, b). Stack: [array, array] → [array of [a_i, b_i] pairs]
pub fn emit_zip(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "zip");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(2, line);
}

/// sum(iterable). Stack: [array] → [number]
pub fn emit_sum(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "sum");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// min(args...). Stack: [args...] → [min_value]
pub fn emit_min(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("vybe:array", "pymin");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(arg_count, line);
}

/// max(args...). Stack: [args...] → [max_value]
pub fn emit_max(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("vybe:array", "pymax");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(arg_count, line);
}
