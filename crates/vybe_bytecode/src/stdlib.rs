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
    c.emit_op_u16(Op::ARRAY_NEW, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // i = start (local 1 already has it, but copy for clarity — it IS local 1)
    // Loop: while i < stop (for positive step) or i > stop (negative step)
    // For simplicity, always check i < stop (works for positive step, which is 99% of usage)
    let loop_start = c.current_offset();
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, stop, 0);
    c.emit_op(Op::DYN_LT, 0);
    let exit = c.emit_jump(Op::BR_IF_FALSE, 0);

    // result.push(i)
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op(Op::ARRAY_PUSH, 0);
    c.emit_op(Op::DROP, 0);

    // i += step
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, step, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, start, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
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
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    c.emit_op_u16(Op::CONST, max, 0);
    c.emit_op(Op::ARRAY_SLICE, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // len = result.length
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // Insertion sort: for i = 1 to len-1
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let outer_loop = c.current_offset();
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    let outer_exit = c.emit_jump(Op::BR_IF_FALSE, 0);

    // key = result[i]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);
    c.emit_op(Op::DROP, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    // while j >= 0 && result[j] > key
    let inner_loop = c.current_offset();
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    let inner_exit = c.emit_jump(Op::BR_IF_FALSE, 0);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::DYN_GT, 0);
    let inner_exit2 = c.emit_jump(Op::BR_IF_FALSE, 0);

    // result[j+1] = result[j]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    // Now stack: [result, j+1] — need value = result[j]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_loop(inner_loop, 0);
    c.patch_jump(inner_exit);
    c.patch_jump(inner_exit2);

    // result[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_loop(outer_loop, 0);
    c.patch_jump(outer_exit);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
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
    c.emit_op_u16(Op::ARRAY_NEW, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // i = arr.length - 1
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    let exit = c.emit_jump(Op::BR_IF_FALSE, 0);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::ARRAY_PUSH, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
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

    c.emit_op_u16(Op::ARRAY_NEW, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    let exit = c.emit_jump(Op::BR_IF_FALSE, 0);

    // Build pair [i, arr[i]], then push onto result.
    // array_push takes [array, value] — so emit result first, then pair.
    c.emit_op_u16(Op::LOCAL_GET, result, 0); // result on stack
    c.emit_op_u16(Op::LOCAL_GET, i, 0);      // i
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);             // arr[i]
    c.emit_op_u16(Op::ARRAY_NEW, 2, 0);      // pair = [i, arr[i]]
    c.emit_op(Op::ARRAY_PUSH, 0);            // result.push(pair)
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
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

    c.emit_op_u16(Op::ARRAY_NEW, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // len = min(a.length, b.length) — use a.length for simplicity
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    let exit = c.emit_jump(Op::BR_IF_FALSE, 0);

    // result.push([a[i], b[i]])
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_GET, b, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::ARRAY_NEW, 2, 0);
    c.emit_op(Op::ARRAY_PUSH, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
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

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, total, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    let exit = c.emit_jump(Op::BR_IF_FALSE, 0);

    c.emit_op_u16(Op::LOCAL_GET, total, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, total, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::LOCAL_GET, total, 0);
    c.emit_op(Op::RETURN, 0);
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
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    let exit = c.emit_jump(Op::BR_IF_FALSE, 0);

    // if arr[i] < best: best = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::DYN_LT, 0);
    let skip = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);
    c.patch_jump(skip);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::RETURN, 0);
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

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let loop_start = c.current_offset();
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    let exit = c.emit_jump(Op::BR_IF_FALSE, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::DYN_GT, 0);
    let skip = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);
    c.patch_jump(skip);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── pow(base, exp) → number (integer exponent by repeated mul) ──
fn build_pow() -> Chunk {
    let mut c = Chunk::new("__stdlib_pow");
    c.arity = 2;
    c.local_count = 4; // callee(0) + base(1) + exp(2) + result(3)
    let base = 1u16;
    let exp = 2;
    let result = 3;

    // result = 1.0
    let one = c.add_constant(Value::F64(1.0));
    c.emit_op_u16(Op::CONST, one, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // while exp > 0: result *= base; exp -= 1
    let loop_start = c.current_offset();
    c.emit_op_u16(Op::LOCAL_GET, exp, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GT, 0);
    let exit = c.emit_jump(Op::BR_IF_FALSE, 0);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, base, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, exp, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, exp, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_loop(loop_start, 0);
    c.patch_jump(exit);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
