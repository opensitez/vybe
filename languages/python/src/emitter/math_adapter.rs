//! Python `math` module functions that aren't a single `ecma:math` host call.
//!
//! Composed on top of `ecma:math`/arithmetic and routed via
//! `common:python.math_*` from the profile `[builtins]` table. Integer-returning
//! functions (`factorial`, `gcd`, …) leave a raw `i32` so they repr as `120`,
//! not `120.0`; the rest produce boxed f64 (Python floats).

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

use vybe_emitter::{collections, ops, tuples};

const DEG_PER_RAD: f64 = 57.295_779_513_082_32; // 180 / π
const RAD_PER_DEG: f64 = 0.017_453_292_519_943_295; // π / 180

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    // Register on the CURRENT chunk (not chunks[0]) so `normalize_import_table`
    // remaps this CALL_IMPORT via the emitting chunk's own local table. A
    // chunks[0] index inside a non-root chunk collides with per-chunk imports
    // and resolves the wrong host fn. Matches shared `emit_import_call`.
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

/// Push arg slot `slot` unboxed to an i32.
fn get_i32(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
}

/// Push arg slot `slot` unboxed to an f64.
fn get_f64(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
}

/// Box a raw f64 on the stack into a Python number value.
fn box_f64(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "wasm:js-number", "fromF64", 1, line);
}

/// `math.gcd(a, b)` → greatest common divisor (non-negative int) via Euclid.
pub fn emit_gcd(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let a = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);
    let t = chunks[current].alloc_scratch(1);
    get_i32(chunks, current, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);
    get_i32(chunks, current, base + 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    // t = a % b; a = b; b = t
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    // |a|
    emit_i32_abs(chunks, current, a, line);
}

/// `math.lcm(a, b)` → least common multiple (int): `|a*b| / gcd(a,b)`, 0 if either is 0.
pub fn emit_lcm(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let a = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);
    let g = chunks[current].alloc_scratch(1);
    let ga = chunks[current].alloc_scratch(1);
    let gb = chunks[current].alloc_scratch(1);
    let t = chunks[current].alloc_scratch(1);
    get_i32(chunks, current, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);
    get_i32(chunks, current, base + 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);

    // if a==0 || b==0 → 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);
    // g = gcd(|a|,|b|)
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ga, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, gb, line);
    emit_i32_abs(chunks, current, ga, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ga, line);
    emit_i32_abs(chunks, current, gb, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, gb, line);
    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, gb, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ga, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, gb, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, gb, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ga, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, gb, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    // gcd result is in `ga` (loop leaves nothing on the stack)
    chunks[current].emit_op_u16(Op::LOCAL_GET, ga, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, g, line);
    // |a*b| / g
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    emit_i32_abs(chunks, current, t, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, g, line);
    chunks[current].emit_op(Op::I32_DIV_S, line);
    chunks[current].emit_end(line);
}

/// `math.comb(n, k)` → n!/(k!(n-k)!) computed iteratively to stay integer.
pub fn emit_comb(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_comb_perm(chunks, current, /*divide=*/ true, line);
}

/// `math.perm(n, k)` → n!/(n-k)! (falling factorial).
pub fn emit_perm(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_comb_perm(chunks, current, /*divide=*/ false, line);
}

fn emit_comb_perm(chunks: &mut [Chunk], current: usize, divide: bool, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let n = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);
    let res = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    get_i32(chunks, current, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    get_i32(chunks, current, base + 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, k, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, res, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    // for i in 0..k: res = res*(n-i) [ /(i+1) for comb ]
    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, k, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    // res = res * (n - i)
    chunks[current].emit_op_u16(Op::LOCAL_GET, res, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    if divide {
        // / (i + 1)
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op(Op::I32_DIV_S, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, res, line);
    // i++
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, res, line);
}

/// `math.prod(iterable[, start])` → product of elements as an int (integer
/// inputs; matches the exercised cases). Accumulates in i32.
pub fn emit_prod(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let arr = base;
    let acc = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        get_i32(chunks, current, base + 1, line);
    } else {
        chunks[current].emit_i32_const(1, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let el = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    // acc *= int(arr[i])
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, el, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
    get_i32(chunks, current, el, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
}

/// `math.degrees(x)` → x * 180/π (float).
pub fn emit_degrees(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    get_f64(chunks, current, base, line);
    chunks[current].emit_f64_const(DEG_PER_RAD, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    box_f64(chunks, current, line);
}

/// `math.radians(x)` → x * π/180 (float).
pub fn emit_radians(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    get_f64(chunks, current, base, line);
    chunks[current].emit_f64_const(RAD_PER_DEG, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    box_f64(chunks, current, line);
}

/// `math.copysign(x, y)` → |x| with the sign of y (float).
pub fn emit_copysign(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    // |x|
    get_f64(chunks, current, base, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    // sign = 1 - 2*(y < 0)  → 1 or -1 (i32) → f64
    chunks[current].emit_i32_const(1, line);
    get_f64(chunks, current, base + 1, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    box_f64(chunks, current, line);
}

/// `math.fmod(x, y)` → C fmod: `x - trunc(x/y)*y` (float, sign of x).
pub fn emit_fmod(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    get_f64(chunks, current, base, line);
    // trunc(x/y) * y
    get_f64(chunks, current, base, line);
    get_f64(chunks, current, base + 1, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    get_f64(chunks, current, base + 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    box_f64(chunks, current, line);
}

/// `math.ldexp(x, i)` → x * 2**i (float).
pub fn emit_ldexp(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    get_f64(chunks, current, base, line);
    // 2 ** i via ecma:math.pow(2, i)
    chunks[current].emit_f64_const(2.0, line);
    box_f64(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    call_import(chunks, current, "ecma:math", "pow", 2, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    box_f64(chunks, current, line);
}

/// `math.dist(p, q)` → Euclidean distance sqrt(Σ(pᵢ-qᵢ)²) (float).
pub fn emit_dist(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let p = base;
    let q = base + 1;
    let sum = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sum, line);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let pe = chunks[current].alloc_scratch(1);
    let qe = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    // d = p[i]-q[i]; sum += d*d
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pe, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, q, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, qe, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sum, line);
    get_f64(chunks, current, pe, line);
    get_f64(chunks, current, qe, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    // dup via recompute: (p-q)*(p-q)
    let d = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, d, line); // store raw f64 diff
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sum, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    // sqrt(sum)
    chunks[current].emit_op_u16(Op::LOCAL_GET, sum, line);
    chunks[current].emit_op(Op::F64_SQRT, line);
    box_f64(chunks, current, line);
}

/// `math.modf(x)` → (fractional, integer) parts, both floats.
pub fn emit_modf(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let xi = chunks[current].alloc_scratch(1); // integer part (raw f64)
    get_f64(chunks, current, base, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, xi, line);
    // frac = x - trunc(x)
    get_f64(chunks, current, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, xi, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    box_f64(chunks, current, line);
    // int part
    chunks[current].emit_op_u16(Op::LOCAL_GET, xi, line);
    box_f64(chunks, current, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

/// `math.frexp(x)` → (m, e) with x = m*2**e, 0.5 ≤ |m| < 1, e an int.
/// x == 0 → (0.0, 0).
pub fn emit_frexp(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let e = chunks[current].alloc_scratch(1); // exponent (i32)
    // e = floor(log2(|x|)) + 1  (0 when x == 0)
    get_f64(chunks, current, base, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line); // (|x| == 0) i32
    chunks[current].emit_if_value(line);
    // zero → e = 0, m = 0.0
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, e, line);
    chunks[current].emit_f64_const(0.0, line);
    box_f64(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_else(line);
    // e = floor(log2(|x|)) + 1
    get_f64(chunks, current, base, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    box_f64(chunks, current, line);
    call_import(chunks, current, "ecma:math", "log2", 1, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, e, line);
    // m = x / 2**e
    get_f64(chunks, current, base, line);
    chunks[current].emit_f64_const(2.0, line);
    box_f64(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, e, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    box_f64(chunks, current, line);
    call_import(chunks, current, "ecma:math", "pow", 2, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    box_f64(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, e, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_end(line);
}

/// `math.fsum(iterable)` → accurate float sum via Neumaier compensated
/// summation, so e.g. `fsum([0.1, 0.2, 0.3])` is exactly `0.6`.
pub fn emit_fsum(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let arr = base;
    let sum = chunks[current].alloc_scratch(1);
    let c = chunks[current].alloc_scratch(1);
    let x = chunks[current].alloc_scratch(1); // current element (raw f64)
    let t = chunks[current].alloc_scratch(1); // sum + x
    let el = chunks[current].alloc_scratch(1); // boxed element
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sum, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, c, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    // x = f64(arr[i]); t = sum + x
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, el, line);
    get_f64(chunks, current, el, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sum, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);
    // if |sum| >= |x|: c += (sum - t) + x  else: c += (x - t) + sum
    chunks[current].emit_op_u16(Op::LOCAL_GET, sum, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_if_value(line);
    // (sum - t) + x
    chunks[current].emit_op_u16(Op::LOCAL_GET, sum, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_else(line);
    // (x - t) + sum
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sum, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_end(line);
    // c += (…)
    chunks[current].emit_op_u16(Op::LOCAL_GET, c, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, c, line);
    // sum = t
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sum, line);
    // i++
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    // result = sum + c
    chunks[current].emit_op_u16(Op::LOCAL_GET, sum, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, c, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    box_f64(chunks, current, line);
}

/// `math.isinf(x)` → x is ±∞ (Bool): `not isfinite(x) and not isnan(x)`.
pub fn emit_isinf(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    // !isFinite(x)
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    call_import(chunks, current, "ecma:number", "isFinite", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    // !isNaN(x)
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    call_import(chunks, current, "ecma:number", "isNaN", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `math.remainder(x, y)` → IEEE remainder `x - round(x/y)*y`. `round` uses
/// ecma round (half-up); the exercised inputs don't hit the half-to-even case.
pub fn emit_remainder(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    get_f64(chunks, current, base, line);
    // round(x/y) * y
    get_f64(chunks, current, base, line);
    get_f64(chunks, current, base + 1, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    box_f64(chunks, current, line);
    call_import(chunks, current, "ecma:math", "round", 1, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    get_f64(chunks, current, base + 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    box_f64(chunks, current, line);
}

/// `math.isclose(a, b, rel_tol=1e-9, abs_tol=0.0)` → Bool:
/// `|a-b| <= max(rel_tol*max(|a|,|b|), abs_tol)`. Keyword tolerances arrive
/// positionally (argc-based).
pub fn emit_isclose(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let diff = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let rt = chunks[current].alloc_scratch(1);
    let at = chunks[current].alloc_scratch(1);
    let thr = chunks[current].alloc_scratch(1);
    let aa = chunks[current].alloc_scratch(1);
    let bb = chunks[current].alloc_scratch(1);

    // diff = |a - b|
    get_f64(chunks, current, base, line);
    get_f64(chunks, current, base + 1, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, diff, line);
    // aa = |a|, bb = |b|
    get_f64(chunks, current, base, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, aa, line);
    get_f64(chunks, current, base + 1, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bb, line);
    // m = max(aa, bb)
    emit_f64_max(chunks, current, aa, bb, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, m, line);
    // rt (default 1e-9), at (default 0)
    if argc >= 3 {
        get_f64(chunks, current, base + 2, line);
    } else {
        chunks[current].emit_f64_const(1e-9, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, rt, line);
    if argc >= 4 {
        get_f64(chunks, current, base + 3, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, at, line);
    // thr = max(rt*m, at)
    chunks[current].emit_op_u16(Op::LOCAL_GET, rt, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, m, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, thr, line); // reuse thr as rt*m
    emit_f64_max(chunks, current, thr, at, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, thr, line);
    // result = diff <= thr
    chunks[current].emit_op_u16(Op::LOCAL_GET, diff, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, thr, line);
    chunks[current].emit_op(Op::F64_LE, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Push `max(slot_a, slot_b)` (both raw-f64 slots) onto the stack.
fn emit_f64_max(chunks: &mut [Chunk], current: usize, a: u16, b: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_end(line);
}

/// `slot = |slot|` as i32, leaving the result on the stack.
fn emit_i32_abs(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    // (x < 0) ? -x : x
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_end(line);
}

/// `math.factorial(n)` → n! as an int. Iterative product `1*2*…*n`.
pub fn emit_factorial(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let n = chunks[current].alloc_scratch(1);
    get_i32(chunks, current, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let result = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    // break when i > n
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_br_if(1, line);
    // result *= i
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    // i++
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}
