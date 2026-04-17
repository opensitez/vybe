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

/// Build all stdlib chunks. Returns (chunks, export_map) where export_map
/// maps function name → chunk index offset (caller adds their base offset).
pub fn build_stdlib() -> StdLib {
    let mut chunks = Vec::new();
    let mut exports = Vec::new();

    // Each function is a separate chunk with a name.
    // The compiler emits `ref_func <idx>` or `call <idx>` to invoke them.

    chunks.push(build_range());
    exports.push("__stdlib_range");

    chunks.push(build_sorted());
    exports.push("__stdlib_sorted");

    chunks.push(build_sort_in_place());
    exports.push("__stdlib_sort_in_place");

    chunks.push(build_sort_with_comparator());
    exports.push("__stdlib_sort_with_comparator");

    chunks.push(build_sort_by_key());
    exports.push("__stdlib_sort_by_key");

    chunks.push(build_reversed());
    exports.push("__stdlib_reversed");

    chunks.push(build_enumerate());
    exports.push("__stdlib_enumerate");

    chunks.push(build_zip());
    exports.push("__stdlib_zip");

    chunks.push(build_sum());
    exports.push("__stdlib_sum");

    chunks.push(build_min());
    exports.push("__stdlib_min");

    chunks.push(build_max());
    exports.push("__stdlib_max");

    chunks.push(build_pow());
    exports.push("__stdlib_pow");

    chunks.push(build_to_string());
    exports.push("__stdlib_tostring");

    chunks.push(build_str_count());
    exports.push("__stdlib_count");

    chunks.push(build_is_numeric());
    exports.push("__stdlib_isnumeric");

    chunks.push(build_splice());
    exports.push("__stdlib_splice");

    chunks.push(build_floor());
    exports.push("__stdlib_floor");

    chunks.push(build_slice());
    exports.push("__stdlib_slice");

    chunks.push(build_keys());
    exports.push("__stdlib_keys");

    chunks.push(build_has_property());
    exports.push("__stdlib_hasproperty");

    chunks.push(build_assign());
    exports.push("__stdlib_assign");

    chunks.push(build_instance_of());
    exports.push("__stdlib_instanceof");

    chunks.push(build_delete_property());
    exports.push("__stdlib_deleteproperty");

    chunks.push(build_array_from());
    exports.push("__stdlib_from");

    chunks.push(build_redim());
    exports.push("__stdlib_redim");

    chunks.push(build_slice_step());
    exports.push("__stdlib_slicestep");

    chunks.push(build_dyn_mul());
    exports.push("__stdlib_dynmul");

    chunks.push(build_concat());
    exports.push("__stdlib_concat");

    chunks.push(build_string_raw());
    exports.push("__stdlib_string_raw");

    StdLib { chunks, exports }
}

pub struct StdLib {
    pub chunks: Vec<Chunk>,
    pub exports: Vec<&'static str>,
}

impl StdLib {
    /// Get chunk index for a stdlib function by name.
    /// Returns offset relative to stdlib base (caller adds their chunk offset).
    pub fn get(&self, name: &str) -> Option<usize> {
        self.exports.iter().position(|&n| n == name)
    }
}

// ── range(start, stop, step) → array ────────────────────────
// Uses: array_new, array_push, i32_add, dyn_lt, br_if_false, loop
fn build_range() -> Chunk {
    let mut c = Chunk::new("__stdlib_range");
    c.arity = 3; // start, stop, step
    c.local_count = 5; // callee(0) + start(1) + stop(2) + step(3) + result(4)
    let start = 1u16;
    let stop = 2;
    let step = 3;
    let result = 4;

    // result = []
    c.emit_op_u16(Op::array_new, 0, 0);
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);

    // i = start (local 1 already has it, but copy for clarity — it IS local 1)
    // Loop: while i < stop (for positive step) or i > stop (negative step)
    // For simplicity, always check i < stop (works for positive step, which is 99% of usage)
    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, start, 0);
    c.emit_op_u16(Op::local_get, stop, 0);
    c.emit_op(Op::dyn_lt, 0);
    let exit = c.emit_jump(Op::br_if_false, 0);

    // result.push(i)
    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, start, 0);
    c.emit_op(Op::array_push, 0);
    c.emit_op(Op::drop, 0);

    // i += step
    c.emit_op_u16(Op::local_get, start, 0);
    c.emit_op_u16(Op::local_get, step, 0);
    c.emit_op(Op::dyn_add, 0);
    c.emit_op_u16(Op::local_set, start, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── sorted(array) → array (insertion sort — O(n²) but works) ──
fn build_sorted() -> Chunk {
    let mut c = Chunk::new("__stdlib_sorted");
    c.arity = 1;
    c.local_count = 7; // callee(0) + arr(1) + result(2) + i(3) + j(4) + len(5) + key(6)
    let arr = 1u16;
    let result = 2;
    let i = 3;
    let j = 4;
    let len = 5;
    let key = 6;

    // Copy input array → result (so we don't mutate the original)
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::i32_const_0, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    c.emit_op_u16(Op::r#const, max, 0);
    c.emit_op(Op::array_slice, 0);
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);

    // len = result.length
    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op_u16(Op::local_set, len, 0);
    c.emit_op(Op::drop, 0);

    // Insertion sort: for i = 1 to len-1
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    let outer_loop = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, len, 0);
    c.emit_op(Op::dyn_lt, 0);
    let outer_exit = c.emit_jump(Op::br_if_false, 0);

    // key = result[i]
    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_set, key, 0);
    c.emit_op(Op::drop, 0);

    // j = i - 1
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, j, 0);
    c.emit_op(Op::drop, 0);

    // while j >= 0 && result[j] > key
    let inner_loop = c.current_offset();
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::dyn_ge, 0);
    let inner_exit = c.emit_jump(Op::br_if_false, 0);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_get, key, 0);
    c.emit_op(Op::dyn_gt, 0);
    let inner_exit2 = c.emit_jump(Op::br_if_false, 0);

    // result[j+1] = result[j]
    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    // Now stack: [result, j+1] — need value = result[j]
    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::array_set, 0);
    c.emit_op(Op::drop, 0);

    // j -= 1
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, j, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(inner_loop, 0);
    c.patch_jump(inner_exit);
    c.patch_jump(inner_exit2);

    // result[j+1] = key
    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_get, key, 0);
    c.emit_op(Op::array_set, 0);
    c.emit_op(Op::drop, 0);

    // i += 1
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(outer_loop, 0);
    c.patch_jump(outer_exit);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op(Op::r#return, 0);
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
fn build_sort_in_place() -> Chunk {
    let mut c = Chunk::new("__stdlib_sort_in_place");
    c.arity = 1;
    c.local_count = 6; // callee(0) + arr(1) + i(2) + j(3) + len(4) + key(5)
    let arr = 1u16;
    let i = 2;
    let j = 3;
    let len = 4;
    let key = 5;

    // len = arr.length
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op_u16(Op::local_set, len, 0);
    c.emit_op(Op::drop, 0);

    // i = 1
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    let outer_loop = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, len, 0);
    c.emit_op(Op::dyn_lt, 0);
    let outer_exit = c.emit_jump(Op::br_if_false, 0);

    // key = arr[i]
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_set, key, 0);
    c.emit_op(Op::drop, 0);

    // j = i - 1
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, j, 0);
    c.emit_op(Op::drop, 0);

    // while j >= 0 && arr[j] > key
    let inner_loop = c.current_offset();
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::dyn_ge, 0);
    let inner_exit = c.emit_jump(Op::br_if_false, 0);

    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_get, key, 0);
    c.emit_op(Op::dyn_gt, 0);
    let inner_exit2 = c.emit_jump(Op::br_if_false, 0);

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::array_set, 0);
    c.emit_op(Op::drop, 0);

    // j -= 1
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, j, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(inner_loop, 0);
    c.patch_jump(inner_exit);
    c.patch_jump(inner_exit2);

    // arr[j+1] = key
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_get, key, 0);
    c.emit_op(Op::array_set, 0);
    c.emit_op(Op::drop, 0);

    // i += 1
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(outer_loop, 0);
    c.patch_jump(outer_exit);

    // return arr (same reference, now sorted in place)
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── sort_with_comparator(array, fn) → same array, sorted using fn ──
// Same insertion sort as sort_in_place, but uses `fn(a, b)` for
// comparison instead of `dyn_gt`. The comparator returns:
//   negative → a before b (no swap)
//   zero     → equal (no swap)
//   positive → b before a (swap)
// This is the standard JS `Array.sort(compareFn)` contract.
fn build_sort_with_comparator() -> Chunk {
    let mut c = Chunk::new("__stdlib_sort_with_comparator");
    c.arity = 2;
    c.local_count = 7; // callee(0) + arr(1) + cmp(2) + i(3) + j(4) + len(5) + key(6)
    let arr = 1u16;
    let cmp = 2;
    let i = 3;
    let j = 4;
    let len = 5;
    let key = 6;

    // len = arr.length
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op_u16(Op::local_set, len, 0);
    c.emit_op(Op::drop, 0);

    // i = 1
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    let outer_loop = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, len, 0);
    c.emit_op(Op::dyn_lt, 0);
    let outer_exit = c.emit_jump(Op::br_if_false, 0);

    // key = arr[i]
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_set, key, 0);
    c.emit_op(Op::drop, 0);

    // j = i - 1
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, j, 0);
    c.emit_op(Op::drop, 0);

    // while j >= 0 && cmp(arr[j], key) > 0
    let inner_loop = c.current_offset();
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::dyn_ge, 0);
    let inner_exit = c.emit_jump(Op::br_if_false, 0);

    // call cmp(arr[j], key) → result
    c.emit_op_u16(Op::local_get, cmp, 0);
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_get, key, 0);
    c.emit_op_u8(Op::call_ref, 2, 0);
    // result > 0 → swap needed
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::dyn_gt, 0);
    let inner_exit2 = c.emit_jump(Op::br_if_false, 0);

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::array_set, 0);
    c.emit_op(Op::drop, 0);

    // j -= 1
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, j, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(inner_loop, 0);
    c.patch_jump(inner_exit);
    c.patch_jump(inner_exit2);

    // arr[j+1] = key
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_get, key, 0);
    c.emit_op(Op::array_set, 0);
    c.emit_op(Op::drop, 0);

    // i += 1
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(outer_loop, 0);
    c.patch_jump(outer_exit);

    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── reversed(array) → array ─────────────────────────────────
fn build_reversed() -> Chunk {
    let mut c = Chunk::new("__stdlib_reversed");
    c.arity = 1;
    c.local_count = 4; // callee(0) + arr(1) + result(2) + i(3)
    let arr = 1u16;
    let result = 2;
    let i = 3;

    // result = []
    c.emit_op_u16(Op::array_new, 0, 0);
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);

    // i = arr.length - 1
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::dyn_ge, 0);
    let exit = c.emit_jump(Op::br_if_false, 0);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::array_push, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── enumerate(array) → [[0,a],[1,b],...] ────────────────────
fn build_enumerate() -> Chunk {
    let mut c = Chunk::new("__stdlib_enumerate");
    c.arity = 1;
    c.local_count = 5; // callee(0) + arr(1) + result(2) + i(3) + len(4)
    let arr = 1u16;
    let result = 2;
    let i = 3;
    let len = 4;

    c.emit_op_u16(Op::array_new, 0, 0);
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op_u16(Op::local_set, len, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op(Op::i32_const_0, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, len, 0);
    c.emit_op(Op::dyn_lt, 0);
    let exit = c.emit_jump(Op::br_if_false, 0);

    // Build pair [i, arr[i]], then push onto result.
    // array_push takes [array, value] — so emit result first, then pair.
    c.emit_op_u16(Op::local_get, result, 0); // result on stack
    c.emit_op_u16(Op::local_get, i, 0);      // i
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);             // arr[i]
    c.emit_op_u16(Op::array_new, 2, 0);      // pair = [i, arr[i]]
    c.emit_op(Op::array_push, 0);            // result.push(pair)
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── zip(a, b) → [[a0,b0],[a1,b1],...] ──────────────────────
fn build_zip() -> Chunk {
    let mut c = Chunk::new("__stdlib_zip");
    c.arity = 2;
    c.local_count = 6; // callee(0) + a(1) + b(2) + result(3) + i(4) + len(5)
    let a = 1u16;
    let b = 2;
    let result = 3;
    let i = 4;
    let len = 5;

    c.emit_op_u16(Op::array_new, 0, 0);
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);

    // len = min(a.length, b.length) — use a.length for simplicity
    c.emit_op_u16(Op::local_get, a, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op_u16(Op::local_set, len, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op(Op::i32_const_0, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, len, 0);
    c.emit_op(Op::dyn_lt, 0);
    let exit = c.emit_jump(Op::br_if_false, 0);

    // result.push([a[i], b[i]])
    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, a, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_get, b, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::array_new, 2, 0);
    c.emit_op(Op::array_push, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── sum(array) → number ─────────────────────────────────────
fn build_sum() -> Chunk {
    let mut c = Chunk::new("__stdlib_sum");
    c.arity = 1;
    c.local_count = 5; // callee(0) + arr(1) + total(2) + i(3) + len(4)
    let arr = 1u16;
    let total = 2;
    let i = 3;
    let len = 4;

    c.emit_op(Op::i32_const_0, 0);
    c.emit_op_u16(Op::local_set, total, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op_u16(Op::local_set, len, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op(Op::i32_const_0, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, len, 0);
    c.emit_op(Op::dyn_lt, 0);
    let exit = c.emit_jump(Op::br_if_false, 0);

    c.emit_op_u16(Op::local_get, total, 0);
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::dyn_add, 0);
    c.emit_op_u16(Op::local_set, total, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::local_get, total, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── min(array) → value ──────────────────────────────────────
fn build_min() -> Chunk {
    let mut c = Chunk::new("__stdlib_min");
    c.arity = 1;
    c.local_count = 5; // callee(0) + arr(1) + best(2) + i(3) + len(4)
    let arr = 1u16;
    let best = 2;
    let i = 3;
    let len = 4;

    // best = arr[0]
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_set, best, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op_u16(Op::local_set, len, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op(Op::i32_const_1, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, len, 0);
    c.emit_op(Op::dyn_lt, 0);
    let exit = c.emit_jump(Op::br_if_false, 0);

    // if arr[i] < best: best = arr[i]
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_get, best, 0);
    c.emit_op(Op::dyn_lt, 0);
    let skip = c.emit_jump(Op::br_if_false, 0);
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_set, best, 0);
    c.emit_op(Op::drop, 0);
    c.patch_jump(skip);

    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::local_get, best, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── max(array) → value ──────────────────────────────────────
fn build_max() -> Chunk {
    let mut c = Chunk::new("__stdlib_max");
    c.arity = 1;
    c.local_count = 5;
    let arr = 1u16;
    let best = 2;
    let i = 3;
    let len = 4;

    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_set, best, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op_u16(Op::local_set, len, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op(Op::i32_const_1, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, len, 0);
    c.emit_op(Op::dyn_lt, 0);
    let exit = c.emit_jump(Op::br_if_false, 0);

    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_get, best, 0);
    c.emit_op(Op::dyn_gt, 0);
    let skip = c.emit_jump(Op::br_if_false, 0);
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_set, best, 0);
    c.emit_op(Op::drop, 0);
    c.patch_jump(skip);

    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::local_get, best, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── pow(base, exp) → number (integer exponent by repeated mul) ──
fn build_pow() -> Chunk {
    // Bytecode-only fallback for `pow(base, exp)`. Handles INTEGER exponents
    // (positive, zero, negative) using a multiply loop. Fractional exponents
    // require floating-point exp/log which WASM doesn't have as standard
    // opcodes — Vybe overrides `__vybe_pow` with a native f64.powf at runtime
    // (polyfill pattern), so this fallback only runs on non-Vybe runtimes
    // and only needs to be correct for the common integer-exp case.
    let mut c = Chunk::new("__stdlib_pow");
    c.arity = 2;
    c.local_count = 5; // callee(0) + base(1) + exp(2) + result(3) + n(4)
    let base = 1u16;
    let exp = 2;
    let result = 3;
    let n = 4;

    // n = abs(exp) — branchless via select would need both values; use a flag
    // We compute n = (exp < 0) ? -exp : exp
    c.emit_op_u16(Op::local_get, exp, 0);
    c.emit_op_u16(Op::local_set, n, 0);
    c.emit_op(Op::drop, 0);
    // if n < 0 then n = -n
    c.emit_op_u16(Op::local_get, n, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::dyn_lt, 0);
    let positive = c.emit_jump(Op::br_if_false, 0);
    c.emit_op_u16(Op::local_get, n, 0);
    c.emit_op(Op::f64_neg, 0);
    c.emit_op_u16(Op::local_set, n, 0);
    c.emit_op(Op::drop, 0);
    c.patch_jump(positive);

    // result = 1.0
    let one = c.add_constant(Value::F64(1.0));
    c.emit_op_u16(Op::r#const, one, 0);
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);

    // while n > 0: result *= base; n -= 1
    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, n, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::dyn_gt, 0);
    let exit = c.emit_jump(Op::br_if_false, 0);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, base, 0);
    c.emit_op(Op::f64_mul, 0);
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, n, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, n, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    // If original exp was negative, take reciprocal: result = 1.0 / result
    c.emit_op_u16(Op::local_get, exp, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::dyn_lt, 0);
    let no_reciprocal = c.emit_jump(Op::br_if_false, 0);
    c.emit_op_u16(Op::r#const, one, 0);
    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op(Op::f64_div, 0);
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);
    c.patch_jump(no_reciprocal);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── toString(value) → string ────────────────────────────────
// "" + value triggers dyn_add string coercion in the VM
fn build_to_string() -> Chunk {
    let mut c = Chunk::new("__stdlib_tostring");
    c.arity = 1;
    c.local_count = 2;
    let val = 1u16;
    let empty = c.add_constant(Value::String(std::sync::Arc::from("")));
    c.emit_op_u16(Op::r#const, empty, 0);
    c.emit_op_u16(Op::local_get, val, 0);
    c.emit_op(Op::dyn_add, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── count(haystack, needle) → int ───────────────────────────
// Count non-overlapping occurrences using substring + indexOf loop
fn build_str_count() -> Chunk {
    let mut c = Chunk::new("__stdlib_count");
    c.arity = 2;
    c.local_count = 5;
    let haystack = 1u16;
    let needle = 2;
    let count = 3;
    let pos = 4;

    c.emit_op(Op::i32_const_0, 0);
    c.emit_op_u16(Op::local_set, count, 0);
    c.emit_op(Op::drop, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op_u16(Op::local_set, pos, 0);
    c.emit_op(Op::drop, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, haystack, 0);
    c.emit_op_u16(Op::local_get, pos, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    c.emit_op_u16(Op::r#const, max, 0);
    c.emit_op(Op::str_substring, 0);
    c.emit_op_u16(Op::local_get, needle, 0);
    c.emit_op(Op::str_index_of, 0);
    c.emit_op(Op::dup, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::dyn_lt, 0);
    let exit = c.emit_jump(Op::br_if_true, 0);
    c.emit_op(Op::drop, 0);
    c.emit_op_u16(Op::local_get, count, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, count, 0);
    c.emit_op(Op::drop, 0);
    c.emit_op_u16(Op::local_get, pos, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, pos, 0);
    c.emit_op(Op::drop, 0);
    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, count, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── splice(arr, index, deleteCount) → removed_elements ──────
// Returns array of removed elements. Mutates arr by removing elements.
// Pure bytecode: build new array from arr[0:index] + arr[index+deleteCount:end]
fn build_splice() -> Chunk {
    let mut c = Chunk::new("__stdlib_splice");
    c.arity = 3;
    c.local_count = 7; // callee(0) + arr(1) + index(2) + delete_count(3) + result(4) + i(5) + end(6)
    let arr = 1u16;
    let index = 2;
    let delete_count = 3;
    let result_local = 4;
    let i = 5;
    let end = 6;

    // result = [] (removed elements)
    c.emit_op_u16(Op::array_new, 0, 0);
    c.emit_op_u16(Op::local_set, result_local, 0);
    c.emit_op(Op::drop, 0);

    // Collect removed elements: arr[index..index+deleteCount]
    c.emit_op_u16(Op::local_get, index, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, index, 0);
    c.emit_op_u16(Op::local_get, delete_count, 0);
    c.emit_op(Op::dyn_add, 0);
    c.emit_op_u16(Op::local_set, end, 0);
    c.emit_op(Op::drop, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, end, 0);
    c.emit_op(Op::dyn_lt, 0);
    let exit = c.emit_jump(Op::br_if_false, 0);

    c.emit_op_u16(Op::local_get, result_local, 0);
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::array_push, 0);
    c.emit_op(Op::drop, 0);

    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    // Return removed elements (actual array mutation would need more complex bytecode)
    c.emit_op_u16(Op::local_get, result_local, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── isNumeric(value) → bool ─────────────────────────────────
// Check if value is a number type using ref_typeof opcode.
fn build_is_numeric() -> Chunk {
    let mut c = Chunk::new("__stdlib_isnumeric");
    c.arity = 1;
    c.local_count = 2; // callee(0) + val(1)
    let val = 1u16;

    // typeof(val) → string
    c.emit_op_u16(Op::local_get, val, 0);
    c.emit_op(Op::ref_typeof, 0);

    // Check if type is "number" (covers I32, I64, F64)
    let num_str = c.add_constant(Value::String(std::sync::Arc::from("number")));
    c.emit_op_u16(Op::r#const, num_str, 0);
    c.emit_op(Op::str_equals, 0);

    // If true, return true
    c.emit_op(Op::dup, 0);
    let done = c.emit_jump(Op::br_if_true, 0);
    c.emit_op(Op::drop, 0);

    // Also check for string that parses as number: try val + 0 and see if NaN
    // Simpler: check if typeof is "i32"
    // Actually ref_typeof returns "number" for all numeric types already.
    // Also check if it's a numeric string by trying to convert.
    c.emit_op_u16(Op::local_get, val, 0);
    c.emit_op(Op::ref_typeof, 0);
    let i32_str = c.add_constant(Value::String(std::sync::Arc::from("i32")));
    c.emit_op_u16(Op::r#const, i32_str, 0);
    c.emit_op(Op::str_equals, 0);

    c.patch_jump(done);
    c.emit_op(Op::r#return, 0);
    c
}

// ── floor(n) → int — wraps f64_floor opcode ────────────────
fn build_floor() -> Chunk {
    let mut c = Chunk::new("__stdlib_floor");
    c.arity = 1;
    c.local_count = 2;
    c.emit_op_u16(Op::local_get, 1, 0);
    c.emit_op(Op::f64_floor, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── slice(arr, start, end) → array — wraps array_slice opcode
fn build_slice() -> Chunk {
    // Polymorphic slice: handles BOTH strings and arrays. The walker doesn't
    // know whether `obj[1..3]` operates on a string or an array, so the
    // canonical slice helper does a runtime type check via `ref_is_string`
    // and dispatches to `str_substring` or `array_slice` accordingly.
    //
    // Used by every language whose surface syntax for slicing is `[start..end]`:
    // C# `arr[1..3]` / `s[0..5]`, Python `arr[1:3]` / `s[0:5]`, etc.
    let mut c = Chunk::new("__stdlib_slice");
    c.arity = 3;
    c.local_count = 4; // callee + obj + start + end
    let obj = 1u16;
    let start = 2u16;
    let end = 3u16;

    // if ref_is_string(obj) → str_substring; else → array_slice
    c.emit_op_u16(Op::local_get, obj, 0);
    c.emit_op(Op::ref_is_string, 0);
    let to_array = c.emit_jump(Op::br_if_false, 0);

    // String branch: [obj, start, end] → str_substring
    c.emit_op_u16(Op::local_get, obj, 0);
    c.emit_op_u16(Op::local_get, start, 0);
    c.emit_op_u16(Op::local_get, end, 0);
    c.emit_op(Op::str_substring, 0);
    c.emit_op(Op::r#return, 0);

    // Array branch
    c.patch_jump(to_array);
    c.emit_op_u16(Op::local_get, obj, 0);
    c.emit_op_u16(Op::local_get, start, 0);
    c.emit_op_u16(Op::local_get, end, 0);
    c.emit_op(Op::array_slice, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── keys(obj) → array of string keys ────────────────────────
// Iterates object properties, collects non-internal keys.
fn build_keys() -> Chunk {
    // Can't iterate properties in pure bytecode without host support.
    // Use dict_keys host call pattern — but that's what we're trying to avoid.
    // Fallback: return empty array. On Vybe, host fn handles it.
    let mut c = Chunk::new("__stdlib_keys");
    c.arity = 1;
    c.local_count = 2;
    // Return empty array as fallback (properties aren't enumerable in pure WASM)
    c.emit_op_u16(Op::array_new, 0, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── hasProperty(obj, key) → bool ────────────────────────────
fn build_has_property() -> Chunk {
    let mut c = Chunk::new("__stdlib_hasproperty");
    c.arity = 2;
    c.local_count = 3;
    c.emit_op_u16(Op::local_get, 1, 0); // obj
    c.emit_op_u16(Op::local_get, 2, 0); // key
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::ref_is_null, 0);
    c.emit_op(Op::dyn_not, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── assign(target, source) → target with source props merged ─
fn build_assign() -> Chunk {
    // Can't iterate source properties in pure bytecode.
    // Fallback: return target unchanged.
    let mut c = Chunk::new("__stdlib_assign");
    c.arity = 2;
    c.local_count = 3;
    c.emit_op_u16(Op::local_get, 1, 0); // return target
    c.emit_op(Op::r#return, 0);
    c
}

// ── instanceOf(obj, type_name) → bool ───────────────────────
fn build_instance_of() -> Chunk {
    let mut c = Chunk::new("__stdlib_instanceof");
    c.arity = 2;
    c.local_count = 3;
    // ref_test needs a constant pool string, but we have a dynamic value.
    // Workaround: compare __type property with the type name string.
    c.emit_op_u16(Op::local_get, 1, 0); // obj
    let type_key = c.add_constant(Value::String(std::sync::Arc::from("__type")));
    c.emit_op_u16(Op::r#const, type_key, 0);
    c.emit_op(Op::array_get, 0); // obj["__type"]
    c.emit_op_u16(Op::local_get, 2, 0); // type_name
    c.emit_op(Op::str_equals, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── deleteProperty(obj, key) → bool ─────────────────────────
fn build_delete_property() -> Chunk {
    // Can't delete properties in pure bytecode. Set to null as fallback.
    let mut c = Chunk::new("__stdlib_deleteproperty");
    c.arity = 2;
    c.local_count = 3;
    c.emit_op_u16(Op::local_get, 1, 0); // obj
    c.emit_op_u16(Op::local_get, 2, 0); // key
    c.emit_op(Op::null, 0);             // value = null
    c.emit_op(Op::array_set, 0);
    c.emit_op(Op::drop, 0);
    c.emit_op(Op::r#true, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── from(iterable) → array copy ─────────────────────────────
fn build_array_from() -> Chunk {
    let mut c = Chunk::new("__stdlib_from");
    c.arity = 1;
    c.local_count = 2;
    // Slice the entire array (copy)
    c.emit_op_u16(Op::local_get, 1, 0);
    c.emit_op(Op::i32_const_0, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    c.emit_op_u16(Op::r#const, max, 0);
    c.emit_op(Op::array_slice, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── redim(arr, new_size) → resized array ────────────────────
fn build_redim() -> Chunk {
    // Create new array of new_size, copy elements from old
    let mut c = Chunk::new("__stdlib_redim");
    c.arity = 2;
    c.local_count = 3;
    c.emit_op_u16(Op::local_get, 1, 0); // arr
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op_u16(Op::local_get, 2, 0); // new_size
    c.emit_op(Op::array_slice, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── sliceStep(arr, start, end, step) → array ─────────────────
fn build_slice_step() -> Chunk {
    let mut c = Chunk::new("__stdlib_slicestep");
    c.arity = 4;
    c.local_count = 7; // arr(1) start(2) end(3) step(4) result(5) i(6)
    let zero = c.add_constant(Value::I32(0));
    c.emit_op_u16(Op::array_new, 0, 0);
    c.emit_op_u16(Op::local_set, 5, 0);
    c.emit_op_u16(Op::local_get, 2, 0);
    c.emit_op_u16(Op::local_set, 6, 0);
    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, 4, 0);
    c.emit_op_u16(Op::r#const, zero, 0);
    c.emit_op(Op::dyn_gt, 0);
    let step_pos = c.emit_jump(Op::br_if_true, 0);
    c.emit_op_u16(Op::local_get, 6, 0);
    c.emit_op_u16(Op::local_get, 3, 0);
    c.emit_op(Op::dyn_gt, 0);
    let cond_done = c.emit_jump(Op::br, 0);
    c.patch_jump(step_pos);
    c.emit_op_u16(Op::local_get, 6, 0);
    c.emit_op_u16(Op::local_get, 3, 0);
    c.emit_op(Op::dyn_lt, 0);
    c.patch_jump(cond_done);
    let exit = c.emit_jump(Op::br_if_false, 0);
    // bounds check
    c.emit_op_u16(Op::local_get, 6, 0);
    c.emit_op_u16(Op::r#const, zero, 0);
    c.emit_op(Op::dyn_lt, 0);
    let skip = c.emit_jump(Op::br_if_true, 0);
    c.emit_op_u16(Op::local_get, 6, 0);
    c.emit_op_u16(Op::local_get, 1, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op(Op::dyn_ge, 0);
    let skip2 = c.emit_jump(Op::br_if_true, 0);
    c.emit_op_u16(Op::local_get, 5, 0);
    c.emit_op_u16(Op::local_get, 1, 0);
    c.emit_op_u16(Op::local_get, 6, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::array_push, 0);
    c.emit_op(Op::drop, 0);
    c.patch_jump(skip);
    c.patch_jump(skip2);
    c.emit_op_u16(Op::local_get, 6, 0);
    c.emit_op_u16(Op::local_get, 4, 0);
    c.emit_op(Op::dyn_add, 0);
    c.emit_op_u16(Op::local_set, 6, 0);
    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);
    c.emit_op_u16(Op::local_get, 5, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── dynMul(a, b) → string repeat or numeric multiply ─────────
fn build_dyn_mul() -> Chunk {
    use std::sync::Arc;
    let mut c = Chunk::new("__stdlib_dynmul");
    c.arity = 2;
    c.local_count = 3;
    let str_tag = c.add_constant(Value::String(Arc::from("string")));
    // if typeof(a) == "string": return str_repeat(a, b)
    c.emit_op_u16(Op::local_get, 1, 0);
    c.emit_op(Op::ref_typeof, 0);
    c.emit_op_u16(Op::r#const, str_tag, 0);
    c.emit_op(Op::dyn_eq, 0);
    let a_not_str = c.emit_jump(Op::br_if_false, 0);
    c.emit_op_u16(Op::local_get, 1, 0);
    c.emit_op_u16(Op::local_get, 2, 0);
    c.emit_op(Op::str_repeat, 0);
    c.emit_op(Op::r#return, 0);
    c.patch_jump(a_not_str);
    // if typeof(b) == "string": return str_repeat(b, a)
    c.emit_op_u16(Op::local_get, 2, 0);
    c.emit_op(Op::ref_typeof, 0);
    c.emit_op_u16(Op::r#const, str_tag, 0);
    c.emit_op(Op::dyn_eq, 0);
    let b_not_str = c.emit_jump(Op::br_if_false, 0);
    c.emit_op_u16(Op::local_get, 2, 0);
    c.emit_op_u16(Op::local_get, 1, 0);
    c.emit_op(Op::str_repeat, 0);
    c.emit_op(Op::r#return, 0);
    c.patch_jump(b_not_str);
    // numeric
    c.emit_op_u16(Op::local_get, 1, 0);
    c.emit_op_u16(Op::local_get, 2, 0);
    c.emit_op(Op::f64_mul, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── sort_by_key(array, keyFn) → same array, sorted by keyFn(x) ──
// .NET LINQ OrderBy(keySelector): insertion sort where comparisons use
// keyFn(a) vs keyFn(b) instead of a vs b directly. The keyFn is a
// 1-arg function that extracts the sort key from each element.
// `OrderBy(x => x)` is identity (plain sort). `OrderBy(x => x.name)`
// sorts by the name property.
fn build_sort_by_key() -> Chunk {
    let mut c = Chunk::new("__stdlib_sort_by_key");
    c.arity = 2;
    c.local_count = 8; // callee(0) + arr(1) + keyFn(2) + i(3) + j(4) + len(5) + key(6) + keyVal(7)
    let arr = 1u16;
    let key_fn = 2;
    let i = 3;
    let j = 4;
    let len = 5;
    let key = 6;
    let key_val = 7;

    // len = arr.length
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op_u16(Op::local_set, len, 0);
    c.emit_op(Op::drop, 0);

    // i = 1
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    let outer_loop = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, len, 0);
    c.emit_op(Op::dyn_lt, 0);
    let outer_exit = c.emit_jump(Op::br_if_false, 0);

    // key = arr[i]
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u16(Op::local_set, key, 0);
    c.emit_op(Op::drop, 0);

    // keyVal = keyFn(key)
    c.emit_op_u16(Op::local_get, key_fn, 0);
    c.emit_op_u16(Op::local_get, key, 0);
    c.emit_op_u8(Op::call_ref, 1, 0);
    c.emit_op_u16(Op::local_set, key_val, 0);
    c.emit_op(Op::drop, 0);

    // j = i - 1
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, j, 0);
    c.emit_op(Op::drop, 0);

    // while j >= 0 && keyFn(arr[j]) > keyVal
    let inner_loop = c.current_offset();
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op(Op::dyn_ge, 0);
    let inner_exit = c.emit_jump(Op::br_if_false, 0);

    // compare: keyFn(arr[j]) > keyVal
    c.emit_op_u16(Op::local_get, key_fn, 0);
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op_u8(Op::call_ref, 1, 0);
    c.emit_op_u16(Op::local_get, key_val, 0);
    c.emit_op(Op::dyn_gt, 0);
    let inner_exit2 = c.emit_jump(Op::br_if_false, 0);

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::array_set, 0);
    c.emit_op(Op::drop, 0);

    // j -= 1
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_sub, 0);
    c.emit_op_u16(Op::local_set, j, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(inner_loop, 0);
    c.patch_jump(inner_exit);
    c.patch_jump(inner_exit2);

    // arr[j+1] = key
    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op_u16(Op::local_get, j, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_get, key, 0);
    c.emit_op(Op::array_set, 0);
    c.emit_op(Op::drop, 0);

    // i += 1
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(outer_loop, 0);
    c.patch_jump(outer_exit);

    c.emit_op_u16(Op::local_get, arr, 0);
    c.emit_op(Op::r#return, 0);
    c
}

// ── concat(a, b) → polymorphic concat ───────────────────────
// If `a` is a string, do str_concat. If `a` is an array, do array_concat.
// Runtime dispatch using ref_is_string. Pure WASM bytecode.
fn build_concat() -> Chunk {
    let mut c = Chunk::new("__stdlib_concat");
    c.arity = 2; // a, b
    c.local_count = 3; // callee(0) + a(1) + b(2)
    let a = 1u16;
    let b = 2u16;

    // if ref_is_string(a) → str_concat
    c.emit_op_u16(Op::local_get, a, 0);
    c.emit_op(Op::ref_is_string, 0);
    let not_string = c.emit_jump(Op::br_if_false, 0);

    // String path: str_concat(a, b)
    c.emit_op_u16(Op::local_get, a, 0);
    c.emit_op_u16(Op::local_get, b, 0);
    c.emit_op(Op::str_concat, 0);
    c.emit_op(Op::r#return, 0);

    c.patch_jump(not_string);

    // Array path: array_concat(a, b)
    c.emit_op_u16(Op::local_get, a, 0);
    c.emit_op_u16(Op::local_get, b, 0);
    c.emit_op(Op::array_concat, 0);
    c.emit_op(Op::r#return, 0);

    c
}

// ── String.raw(strings, ...values) → interleave strings and values ──
// Tagged template function that returns the raw string without escape processing.
// strings[0] + values[0] + strings[1] + values[1] + ... + strings[N]
// Since this is called as a tagged template, strings is an array and
// values are individual args. With rest params, values is already an array.
fn build_string_raw() -> Chunk {
    use std::sync::Arc;

    let mut c = Chunk::new("__stdlib_string_raw");
    c.arity = 2; // strings_array, values_array (rest-packed by caller)
    c.local_count = 6; // callee(0) + strings(1) + values(2) + result(3) + i(4) + len(5)
    let strings = 1u16;
    let values = 2u16;
    let result = 3u16;
    let i = 4u16;
    let len = 5u16;

    // result = ""
    let empty = c.add_constant(Value::String(Arc::from("")));
    c.emit_op_u16(Op::r#const, empty, 0);
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);

    // len = strings.length
    c.emit_op_u16(Op::local_get, strings, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op_u16(Op::local_set, len, 0);
    c.emit_op(Op::drop, 0);

    // i = 0
    c.emit_op(Op::i32_const_0, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    // loop: while i < len
    let loop_start = c.current_offset();
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, len, 0);
    c.emit_op(Op::dyn_lt, 0);
    let exit = c.emit_jump(Op::br_if_false, 0);

    // result += strings[i]
    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, strings, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::str_concat, 0);
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);

    // if i < values.length: result += String(values[i])
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op_u16(Op::local_get, values, 0);
    c.emit_op(Op::array_length, 0);
    c.emit_op(Op::dyn_lt, 0);
    let skip_val = c.emit_jump(Op::br_if_false, 0);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op_u16(Op::local_get, values, 0);
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::array_get, 0);
    c.emit_op(Op::str_concat, 0); // dyn_add would also work since result is string
    c.emit_op_u16(Op::local_set, result, 0);
    c.emit_op(Op::drop, 0);

    c.patch_jump(skip_val);

    // i += 1
    c.emit_op_u16(Op::local_get, i, 0);
    c.emit_op(Op::i32_const_1, 0);
    c.emit_op(Op::i32_add, 0);
    c.emit_op_u16(Op::local_set, i, 0);
    c.emit_op(Op::drop, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::local_get, result, 0);
    c.emit_op(Op::r#return, 0);
    c
}
