//! WASM Standard Library — pure bytecode implementations of common functions.
//!
//! These are compiled to Chunk objects that get linked into every program.
//! On Vybe VM, the compiler can skip these and use host calls instead (faster).
//! On any other WASM runtime, these provide the same functionality portably.
//!
//! Functions:
//!   range(start, stop, step) → array
//!   sorted(array) → array (insertion sort)
//!   reversed(array) → array
//!   enumerate(array) → array of [i, val]
//!   zip(a, b) → array of [a_i, b_i]
//!   sum(array) → number
//!   min_val(array) → value
//!   max_val(array) → value
//!   to_str(value) → string (via convert import)
//!   pow(base, exp) → number (repeated multiplication fallback)

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

/// Build all stdlib chunks. Each chunk registers any `wasm:js-array.*`
/// imports on the passed `imports` chunk (= user program's
/// `chunks[0]`, the module-level imports section per WASM semantics).
/// Returns the stdlib chunks + their export names, in matching order;
/// caller appends the chunks to its own vec.
pub fn build_stdlib(imports: &mut Chunk) -> StdLib {
    let mut chunks = Vec::new();
    let mut exports = Vec::new();

    chunks.push(build_range(imports));             exports.push("__stdlib_range");
    chunks.push(build_sorted(imports));            exports.push("__stdlib_sorted");
    chunks.push(build_sort_in_place(imports));     exports.push("__stdlib_sort_in_place");
    chunks.push(build_sort_with_comparator(imports)); exports.push("__stdlib_sort_with_comparator");
    chunks.push(build_sort_by_key(imports));       exports.push("__stdlib_sort_by_key");
    chunks.push(build_reversed(imports));          exports.push("__stdlib_reversed");
    chunks.push(build_enumerate(imports));         exports.push("__stdlib_enumerate");
    chunks.push(build_zip(imports));               exports.push("__stdlib_zip");
    chunks.push(build_sum(imports));               exports.push("__stdlib_sum");
    chunks.push(build_min(imports));               exports.push("__stdlib_min");
    chunks.push(build_max(imports));               exports.push("__stdlib_max");
    chunks.push(build_pow(imports));               exports.push("__stdlib_pow");
    chunks.push(build_to_string(imports));         exports.push("__stdlib_tostring");
    chunks.push(build_str_count(imports));         exports.push("__stdlib_count");
    chunks.push(build_is_numeric(imports));        exports.push("__stdlib_isnumeric");
    chunks.push(build_splice(imports));            exports.push("__stdlib_splice");
    chunks.push(build_floor(imports));             exports.push("__stdlib_floor");
    chunks.push(build_slice(imports));             exports.push("__stdlib_slice");
    chunks.push(build_keys(imports));              exports.push("__stdlib_keys");
    chunks.push(build_has_property(imports));      exports.push("__stdlib_hasproperty");
    chunks.push(build_assign(imports));            exports.push("__stdlib_assign");
    chunks.push(build_instance_of(imports));       exports.push("__stdlib_instanceof");
    chunks.push(build_delete_property(imports));   exports.push("__stdlib_deleteproperty");
    chunks.push(build_array_from(imports));        exports.push("__stdlib_from");
    chunks.push(build_redim(imports));             exports.push("__stdlib_redim");
    chunks.push(build_slice_step(imports));        exports.push("__stdlib_slicestep");
    chunks.push(build_dyn_mul(imports));           exports.push("__stdlib_dynmul");
    chunks.push(build_concat(imports));            exports.push("__stdlib_concat");
    chunks.push(build_string_raw(imports));        exports.push("__stdlib_string_raw");
    chunks.push(build_fmod(imports));              exports.push("__stdlib_fmod");

    StdLib { chunks, exports }
}

pub struct StdLib {
    pub chunks: Vec<Chunk>,
    pub exports: Vec<&'static str>,
}

impl StdLib {
    pub fn get(&self, name: &str) -> Option<usize> {
        self.exports.iter().position(|&n| n == name)
    }
}

// ── range(start, stop, step) → array ────────────────────────
// Every dynamic-array op routes through `common::collections::emit_*`
// so the emitted bytecode imports `wasm:js-array.*` — works natively on
// v8, on Vybe (registered handlers), and on plain wasmtime with the
// polyfill module. Raw ARRAY_* opcodes are Vybe-only and have been
// removed.
fn build_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_range");
    c.arity = 3; // start, stop, step
    c.local_count = 4;
    let start = 0u16;
    let stop = 1;
    let step = 2;
    let result = 3;

    // result = []
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, stop, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);

    // result.push(i) — push returns new_length (ECMA-262); drop it.
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, step, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, start, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sorted(array) → array (insertion sort — O(n²) but works) ──
fn build_sorted(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sorted");
    c.arity = 1;
    c.local_count = 6; // arr(0) + result(1) + i(2) + j(3) + len(4) + key(5)
    let arr = 0u16;
    let result = 1;
    let i = 2;
    let j = 3;
    let len = 4;
    let key = 5;

    // Copy input array → result (so we don't mutate the original)
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    c.emit_op_u16(Op::CONST, max, 0);
    crate::emitter::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // len = result.length
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // Insertion sort: for i = 1 to len-1
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = result[i]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);
    c.emit_op(Op::DROP, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    // while j >= 0 && result[j] > key
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // result[j+1] = result[j]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    // Now stack: [result, j+1] — need value = result[j]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0); c.patch_loop(inner_loop_p);
    c.emit_end(0); c.patch_block(inner_block_p);

    // result[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0); c.patch_loop(outer_loop_p);
    c.emit_end(0); c.patch_block(outer_block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sort_in_place(array) → same array, mutated ──────────────
// In-place insertion sort. Used by every language whose surface syntax for
// sorting is in-place: C# `list.Sort()`, VB `list.Sort()`, JS `arr.sort()`,
// Python `list.sort()`, Pascal `Sort(arr)`. The walker normalizes each form
// into a canonical builtin call which routes here through compiler_common.
//
// Insertion sort is O(n²) but small and works on arbitrary value comparisons
// via dyn_gt. Higher-perf algorithms can be added behind the same name later.
fn build_sort_in_place(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sort_in_place");
    c.arity = 1;
    c.local_count = 5; // arr(0) + i(1) + j(2) + len(3) + key(4)
    let arr = 0u16;
    let i = 1;
    let j = 2;
    let len = 3;
    let key = 4;

    // len = arr.length
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // i = 1
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);
    c.emit_op(Op::DROP, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    // while j >= 0 && arr[j] > key
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0); c.patch_loop(inner_loop_p);
    c.emit_end(0); c.patch_block(inner_block_p);

    // arr[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0); c.patch_loop(outer_loop_p);
    c.emit_end(0); c.patch_block(outer_block_p);

    // return arr (same reference, now sorted in place)
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sort_with_comparator(array, fn) → same array, sorted using fn ──
// Same insertion sort as sort_in_place, but uses `fn(a, b)` for
// comparison instead of `dyn_gt`. The comparator returns:
//   negative → a before b (no swap)
//   zero     → equal (no swap)
//   positive → b before a (swap)
// This is the standard JS `Array.sort(compareFn)` contract.
fn build_sort_with_comparator(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sort_with_comparator");
    c.arity = 2;
    c.local_count = 6; // arr(0) + cmp(1) + i(2) + j(3) + len(4) + key(5)
    let arr = 0u16;
    let cmp = 1;
    let i = 2;
    let j = 3;
    let len = 4;
    let key = 5;

    // len = arr.length
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // i = 1
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);
    c.emit_op(Op::DROP, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    // while j >= 0 && cmp(arr[j], key) > 0
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop

    // call cmp(arr[j], key) → result
    c.emit_op_u16(Op::LOCAL_GET, cmp, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op_u8(Op::CALL_REF, 2, 0);
    // result > 0 → swap needed
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0); c.patch_loop(inner_loop_p);
    c.emit_end(0); c.patch_block(inner_block_p);

    // arr[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0); c.patch_loop(outer_loop_p);
    c.emit_end(0); c.patch_block(outer_block_p);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── reversed(array) → array ─────────────────────────────────
fn build_reversed(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_reversed");
    c.arity = 1;
    c.local_count = 3; // arr(0) + result(1) + i(2)
    let arr = 0u16;
    let result = 1;
    let i = 2;

    // result = []
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // i = arr.length - 1
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── enumerate(array) → [[0,a],[1,b],...] ────────────────────
fn build_enumerate(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_enumerate");
    c.arity = 1;
    c.local_count = 4; // arr(0) + result(1) + i(2) + len(3)
    let arr = 0u16;
    let result = 1;
    let i = 2;
    let len = 3;

    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    // Build pair [i, arr[i]], then push onto result.
    // array_push takes [array, value] — so emit result first, then pair.
    c.emit_op_u16(Op::LOCAL_GET, result, 0); // result on stack
    c.emit_op_u16(Op::LOCAL_GET, i, 0);      // i
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);             // arr[i]
    crate::emitter::collections::emit_array_pair_into(imports, &mut c, 0);      // pair = [i, arr[i]]
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);            // result.push(pair)
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── zip(a, b) → [[a0,b0],[a1,b1],...] ──────────────────────
fn build_zip(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_zip");
    c.arity = 2;
    c.local_count = 5; // a(0) + b(1) + result(2) + i(3) + len(4)
    let a = 0u16;
    let b = 1;
    let result = 2;
    let i = 3;
    let len = 4;

    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // len = min(a.length, b.length) — use a.length for simplicity
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    // result.push([a[i], b[i]])
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, b, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_array_pair_into(imports, &mut c, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sum(array) → number ─────────────────────────────────────
fn build_sum(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sum");
    c.arity = 1;
    c.local_count = 4; // arr(0) + total(1) + i(2) + len(3)
    let arr = 0u16;
    let total = 1;
    let i = 2;
    let len = 3;

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, total, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, total, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, total, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, total, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── min(array) → value ──────────────────────────────────────
fn build_min(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_min");
    c.arity = 1;
    c.local_count = 4; // arr(0) + best(1) + i(2) + len(3)
    let arr = 0u16;
    let best = 1;
    let i = 2;
    let len = 3;

    // best = arr[0]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    // if arr[i] < best: best = arr[i]
    // block must wrap ALL condition operands + comparison + body
    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip if NOT less than
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(skip_block_p);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── max(array) → value ──────────────────────────────────────
fn build_max(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_max");
    c.arity = 1;
    c.local_count = 4;
    let arr = 0u16;
    let best = 1;
    let i = 2;
    let len = 3;

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip if NOT greater than
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(skip_block_p);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── pow(base, exp) → number (integer exponent by repeated mul) ──
fn build_pow(imports: &mut Chunk) -> Chunk {
    // Bytecode-only fallback for `pow(base, exp)`. Handles INTEGER exponents
    // (positive, zero, negative) using a multiply loop. Fractional exponents
    // require floating-point exp/log which WASM doesn't have as standard
    // opcodes — Vybe overrides `__vybe_pow` with a native f64.powf at runtime
    // (polyfill pattern), so this fallback only runs on non-Vybe runtimes
    // and only needs to be correct for the common integer-exp case.
    let mut c = Chunk::new("__stdlib_pow");
    c.arity = 2;
    c.local_count = 4; // base(0) + exp(1) + result(2) + n(3)
    let base = 0u16;
    let exp = 1;
    let result = 2;
    let n = 3;

    // n = abs(exp) — branchless via select would need both values; use a flag
    // We compute n = (exp < 0) ? -exp : exp
    c.emit_op_u16(Op::LOCAL_GET, exp, 0);
    c.emit_op_u16(Op::LOCAL_SET, n, 0);
    c.emit_op(Op::DROP, 0);
    // if n < 0 then n = -n
    let pos_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip negate if NOT negative
    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op(Op::F64_NEG, 0);
    c.emit_op_u16(Op::LOCAL_SET, n, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(pos_block_p);

    // result = 1.0
    let one = c.add_constant(Value::F64(1.0));
    c.emit_op_u16(Op::CONST, one, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // while n > 0: result *= base; n -= 1
    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, base, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, n, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    // If original exp was negative, take reciprocal: result = 1.0 / result
    let recip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, exp, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip reciprocal if NOT negative
    c.emit_op_u16(Op::CONST, one, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::F64_DIV, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(recip_block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── toString(value) → string ────────────────────────────────
// "" + value triggers dyn_add string coercion in the VM
fn build_to_string(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_tostring");
    c.arity = 1;
    c.local_count = 1;
    let val = 0u16;
    let empty = c.add_constant(Value::String(std::sync::Arc::from("")));
    c.emit_op_u16(Op::CONST, empty, 0);
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── count(haystack, needle) → int ───────────────────────────
// Count non-overlapping occurrences using substring + indexOf loop
fn build_str_count(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_count");
    c.arity = 2;
    c.local_count = 4;
    let haystack = 0u16;
    let needle = 1;
    let count = 2;
    let pos = 3;

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, count, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, pos, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, haystack, 0);
    c.emit_op_u16(Op::LOCAL_GET, pos, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    c.emit_op_u16(Op::CONST, max, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op_u16(Op::LOCAL_GET, needle, 0);
    c.emit_op(Op::STR_INDEX_OF, 0);
    // Save indexOf result to local (don't use DUP — value can't cross block boundary)
    let idx_result = 4u16; // reuse local slot (local_count=4, slot 4 is beyond declared but safe with extra locals)
    c.local_count = 5; // need one more local for idx_result
    c.emit_op_u16(Op::LOCAL_SET, idx_result, 0);
    c.emit_op(Op::DROP, 0);
    // Check if index < 0
    c.emit_op_u16(Op::LOCAL_GET, idx_result, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_br_if(1, 0); // exit loop if index < 0
    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, count, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, pos, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, pos, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── splice(arr, index, deleteCount) → removed_elements ──────
// Returns array of removed elements. Mutates arr by removing elements.
// Pure bytecode: build new array from arr[0:index] + arr[index+deleteCount:end]
fn build_splice(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_splice");
    c.arity = 3;
    c.local_count = 6; // arr(0) + index(1) + delete_count(2) + result(3) + i(4) + end(5)
    let arr = 0u16;
    let index = 1;
    let delete_count = 2;
    let result_local = 3;
    let i = 4;
    let end = 5;

    // result = [] (removed elements)
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result_local, 0);
    c.emit_op(Op::DROP, 0);

    // Collect removed elements: arr[index..index+deleteCount]
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_GET, delete_count, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, end, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, result_local, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    // Return removed elements (actual array mutation would need more complex bytecode)
    c.emit_op_u16(Op::LOCAL_GET, result_local, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── isNumeric(value) → bool ─────────────────────────────────
// Check if value is a number type using ref_typeof opcode.
fn build_is_numeric(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isnumeric");
    c.arity = 1;
    c.local_count = 1; // val(0)
    let val = 0u16;

    // Check if type is "number" (covers I32, I64, F64)
    // Block must wrap ALL values consumed inside it (typeof result + STR_EQUALS + DUP)
    let num_str = c.add_constant(Value::String(std::sync::Arc::from("number")));
    let done_block_p = c.emit_block(0);
    // typeof(val) → string
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, num_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);

    // If true, save and skip second check
    let result_slot = 1u16;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_SET, result_slot, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, result_slot, 0);
    c.emit_br_if(0, 0); // skip to end if already true

    // Also check if typeof is "i32"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    let i32_str = c.add_constant(Value::String(std::sync::Arc::from("i32")));
    c.emit_op_u16(Op::CONST, i32_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op_u16(Op::LOCAL_SET, result_slot, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_end(0); c.patch_block(done_block_p);
    c.emit_op_u16(Op::LOCAL_GET, result_slot, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── floor(n) → int — wraps f64_floor opcode ────────────────
fn build_floor(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_floor");
    c.arity = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::F64_FLOOR, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── slice(arr, start, end) → array — wraps array_slice opcode
fn build_slice(imports: &mut Chunk) -> Chunk {
    // Polymorphic slice: handles BOTH strings and arrays. The walker doesn't
    // know whether `obj[1..3]` operates on a string or an array, so the
    // canonical slice helper does a runtime type check via `ref_is_string`
    // and dispatches to `str_substring` or `array_slice` accordingly.
    //
    // Used by every language whose surface syntax for slicing is `[start..end]`:
    // C# `arr[1..3]` / `s[0..5]`, Python `arr[1:3]` / `s[0:5]`, etc.
    let mut c = Chunk::new("__stdlib_slice");
    c.arity = 3;
    c.local_count = 3; // obj + start + end
    let obj = 0u16;
    let start = 1u16;
    let end = 2u16;

    // if ref_is_string(obj) → str_substring; else → array_slice
    let str_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, obj, 0);
    c.emit_op(Op::REF_IS_STRING, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip string branch if NOT string

    // String branch: [obj, start, end] → str_substring
    c.emit_op_u16(Op::LOCAL_GET, obj, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op(Op::RETURN, 0);

    // Array branch
    c.emit_end(0); c.patch_block(str_block_p);
    c.emit_op_u16(Op::LOCAL_GET, obj, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    crate::emitter::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── keys(obj) → array of string keys ────────────────────────
// Iterates object properties, collects non-internal keys.
fn build_keys(imports: &mut Chunk) -> Chunk {
    // Can't iterate properties in pure bytecode without host support.
    // Use dict_keys host call pattern — but that's what we're trying to avoid.
    // Fallback: return empty array. On Vybe, host fn handles it.
    let mut c = Chunk::new("__stdlib_keys");
    c.arity = 1;
    c.local_count = 1;
    // Return empty array as fallback (properties aren't enumerable in pure WASM)
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── hasProperty(obj, key) → bool ────────────────────────────
fn build_has_property(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_hasproperty");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // obj
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // key
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── assign(target, source) → target with source props merged ─
fn build_assign(imports: &mut Chunk) -> Chunk {
    // Can't iterate source properties in pure bytecode.
    // Fallback: return target unchanged.
    let mut c = Chunk::new("__stdlib_assign");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // return target
    c.emit_op(Op::RETURN, 0);
    c
}

// ── instanceOf(obj, type_name) → bool ───────────────────────
fn build_instance_of(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_instanceof");
    c.arity = 2;
    c.local_count = 2;
    // ref_test needs a constant pool string, but we have a dynamic value.
    // Workaround: compare __type property with the type name string.
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // obj
    let type_key = c.add_constant(Value::String(std::sync::Arc::from("__type")));
    c.emit_op_u16(Op::CONST, type_key, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0); // obj["__type"]
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // type_name
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── deleteProperty(obj, key) → bool ─────────────────────────
fn build_delete_property(imports: &mut Chunk) -> Chunk {
    // Can't delete properties in pure bytecode. Set to null as fallback.
    let mut c = Chunk::new("__stdlib_deleteproperty");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // obj
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // key
    c.emit_op(Op::NULL, 0);             // value = null
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op(Op::TRUE, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── from(iterable) → array copy ─────────────────────────────
fn build_array_from(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_from");
    c.arity = 1;
    c.local_count = 1;
    // Slice the entire array (copy)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    c.emit_op_u16(Op::CONST, max, 0);
    crate::emitter::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── redim(arr, new_size) → resized array ────────────────────
fn build_redim(imports: &mut Chunk) -> Chunk {
    // Create new array of new_size, copy elements from old
    let mut c = Chunk::new("__stdlib_redim");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // arr
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // new_size
    crate::emitter::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sliceStep(arr, start, end, step) → array ─────────────────
fn build_slice_step(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_slicestep");
    c.arity = 4;
    c.local_count = 7; // arr(0) start(1) end(2) step(3) result(4) i(5) cond(6)
    let zero = c.add_constant(Value::I32(0));

    // result = new array
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 4, 0);
    c.emit_op(Op::DROP, 0);
    // i = start
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, 5, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // Compute condition: if step > 0 then i < end else i > end
    // Store in local 6 (cond) to avoid value-on-stack across branches
    let pos_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip positive branch if step <= 0
    // positive step: cond = i < end
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op_u16(Op::LOCAL_SET, 6, 0);
    c.emit_op(Op::DROP, 0);
    let skip_neg_p = c.emit_block(0);
    c.emit_br(1, 0); // skip negative branch (jump past skip_neg block end + neg block)
    c.emit_end(0); c.patch_block(skip_neg_p);
    c.emit_end(0); c.patch_block(pos_block_p);

    // negative step: cond = i > end
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op_u16(Op::LOCAL_SET, 6, 0);
    c.emit_op(Op::DROP, 0);

    // Check condition — exit if false
    c.emit_op_u16(Op::LOCAL_GET, 6, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop (depth 1 = outer block)

    // bounds check: skip push if i < 0 or i >= arr.length
    // Block must wrap the condition values consumed by br_if inside it
    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_br_if(0, 0); // skip push if i < 0
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_br_if(0, 0); // skip push if i >= length
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(skip_block_p);

    // i = i + step
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 5, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── dynMul(a, b) → string repeat or numeric multiply ─────────
fn build_dyn_mul(imports: &mut Chunk) -> Chunk {
    use std::sync::Arc;
    let mut c = Chunk::new("__stdlib_dynmul");
    c.arity = 2;
    c.local_count = 2;
    let str_tag = c.add_constant(Value::String(Arc::from("string")));
    // if typeof(a) == "string": return str_repeat(a, b)
    let a_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, str_tag, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip if a is NOT string
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::STR_REPEAT, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(a_block_p);
    // if typeof(b) == "string": return str_repeat(b, a)
    let b_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, str_tag, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip if b is NOT string
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::STR_REPEAT, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(b_block_p);
    // numeric
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sort_by_key(array, keyFn) → same array, sorted by keyFn(x) ──
// .NET LINQ OrderBy(keySelector): insertion sort where comparisons use
// keyFn(a) vs keyFn(b) instead of a vs b directly. The keyFn is a
// 1-arg function that extracts the sort key from each element.
// `OrderBy(x => x)` is identity (plain sort). `OrderBy(x => x.name)`
// sorts by the name property.
fn build_sort_by_key(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sort_by_key");
    c.arity = 2;
    c.local_count = 7; // arr(0) + keyFn(1) + i(2) + j(3) + len(4) + key(5) + keyVal(6)
    let arr = 0u16;
    let key_fn = 1;
    let i = 2;
    let j = 3;
    let len = 4;
    let key = 5;
    let key_val = 6;

    // len = arr.length
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // i = 1
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);
    c.emit_op(Op::DROP, 0);

    // keyVal = keyFn(key)
    c.emit_op_u16(Op::LOCAL_GET, key_fn, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op_u8(Op::CALL_REF, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, key_val, 0);
    c.emit_op(Op::DROP, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    // while j >= 0 && keyFn(arr[j]) > keyVal
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop

    // compare: keyFn(arr[j]) > keyVal
    c.emit_op_u16(Op::LOCAL_GET, key_fn, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u8(Op::CALL_REF, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, key_val, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0); c.patch_loop(inner_loop_p);
    c.emit_end(0); c.patch_block(inner_block_p);

    // arr[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0); c.patch_loop(outer_loop_p);
    c.emit_end(0); c.patch_block(outer_block_p);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── concat(a, b) → polymorphic concat ───────────────────────
// If `a` is a string, do str_concat. If `a` is an array, do array_concat.
// Runtime dispatch using ref_is_string. Pure WASM bytecode.
fn build_concat(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_concat");
    c.arity = 2; // a, b
    c.local_count = 2; // a(0) + b(1)
    let a = 0u16;
    let b = 1u16;

    // if ref_is_string(a) → str_concat
    let str_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op(Op::REF_IS_STRING, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip string path if NOT string

    // String path: str_concat(a, b)
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op_u16(Op::LOCAL_GET, b, 0);
    c.emit_op(Op::STR_CONCAT, 0);
    c.emit_op(Op::RETURN, 0);

    c.emit_end(0); c.patch_block(str_block_p);

    // Array path: array_concat(a, b)
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op_u16(Op::LOCAL_GET, b, 0);
    crate::emitter::collections::emit_concat_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);

    c
}

// ── String.raw(strings, ...values) → interleave strings and values ──
// Tagged template function that returns the raw string without escape processing.
// strings[0] + values[0] + strings[1] + values[1] + ... + strings[N]
// Since this is called as a tagged template, strings is an array and
// values are individual args. With rest params, values is already an array.
fn build_string_raw(imports: &mut Chunk) -> Chunk {
    use std::sync::Arc;

    let mut c = Chunk::new("__stdlib_string_raw");
    c.arity = 2; // strings_array, values_array (rest-packed by caller)
    c.local_count = 5; // strings(0) + values(1) + result(2) + i(3) + len(4)
    let strings = 0u16;
    let values = 1u16;
    let result = 2u16;
    let i = 3u16;
    let len = 4u16;

    // result = ""
    let empty = c.add_constant(Value::String(Arc::from("")));
    c.emit_op_u16(Op::CONST, empty, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // len = strings.length
    c.emit_op_u16(Op::LOCAL_GET, strings, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // i = 0
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    // loop: while i < len
    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    // result += strings[i]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, strings, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op(Op::STR_CONCAT, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // if i < values.length: result += String(values[i])
    let skip_val_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, values, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip if i >= values.length

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, values, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op(Op::STR_CONCAT, 0); // dyn_add would also work since result is string
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_end(0); c.patch_block(skip_val_p);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── fmod(a, b) → a % b (floating-point remainder) ──────────
// WASM has no f64.rem. Pure bytecode: a - trunc(a/b) * b.
// Host can override __vybe_fmod with native fmod for performance.
fn build_fmod(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_fmod");
    c.arity = 2; // a, b
    c.local_count = 2; // a(0) + b(1)
    let a = 0u16;
    let b = 1u16;

    // result = a - trunc(a / b) * b
    c.emit_op_u16(Op::LOCAL_GET, a, 0);   // a
    c.emit_op_u16(Op::LOCAL_GET, a, 0);   // a
    c.emit_op_u16(Op::LOCAL_GET, b, 0);   // b
    c.emit_op(Op::F64_DIV, 0);            // a / b
    c.emit_op(Op::F64_TRUNC, 0);          // trunc(a / b)
    c.emit_op_u16(Op::LOCAL_GET, b, 0);   // b
    c.emit_op(Op::F64_MUL, 0);            // trunc(a / b) * b
    c.emit_op(Op::F64_SUB, 0);            // a - trunc(a / b) * b
    c.emit_op(Op::RETURN, 0);
    c
}
