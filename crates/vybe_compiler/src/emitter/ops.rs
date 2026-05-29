//! Dynamic-dispatch opcode emitters — spec-compliant replacements for DYN_*.
//!
//! Each function emits the WASM-standard type-dispatch sequence for one of
//! the 10 `DYN_*` VM-internal opcodes, using only:
//!   - Standard WASM opcodes
//!   - `wasm:js-*` host imports (js-string-builtins + js-primitive-builtins proposals)
//!
//! All functions take `(chunk: &mut Chunk, line: u32)`.
//! Imports are added to the same chunk (deduplicated by `add_import`).
//! Local slots are allocated by bumping `chunk.local_count`.
//!
//! ## Branch convention
//!
//! Host functions that return a type-test result push `Value::I32(0 or 1)`.
//! The VM's `BR_IF_FALSE`/`BR_IF_TRUE` check `val.as_bool()`, which only
//! returns true for `Value::Bool(true)`. So after any I32-returning host call,
//! we use `I32_EQZ` + `BR_IF_TRUE` to branch when the result was 0 (false),
//! and simply `BR_IF_TRUE` after `I32_EQZ` to branch when the result was 1.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

// ── helpers ────────────────────────────────────────────────────────────

fn call1(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, line);
    chunk.emit(1, line);
}

fn call2(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, line);
    chunk.emit(2, line);
}

fn alloc_locals(chunk: &mut Chunk, n: u16) -> u16 {
    let base = chunk.local_count;
    chunk.local_count += n;
    base
}

fn save(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line); // peeks, value stays on stack
    chunk.emit_op(Op::DROP, line);                // remove residue
}

fn load(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn i32_const(chunk: &mut Chunk, v: i32, line: u32) {
    let k = chunk.add_constant(Value::I32(v));
    chunk.emit_op_u16(Op::CONST, k, line);
}

fn i64_const(chunk: &mut Chunk, v: i64, line: u32) {
    let k = chunk.add_constant(Value::I64(v));
    chunk.emit_op_u16(Op::CONST, k, line);
}

fn f64_const(chunk: &mut Chunk, v: f64, line: u32) {
    let k = chunk.add_constant(Value::F64(v));
    chunk.emit_op_u16(Op::CONST, k, line);
}

/// Call `test_fn` with the value at `slot`, then branch to `skip_label`
/// if the test returned 0 (not this type).
/// Stack before: []  Stack after: []  (test result consumed by branch)
fn test_and_skip_if_not(chunk: &mut Chunk, slot: u16, test_fn: u16, line: u32) -> usize {
    load(chunk, slot, line);
    call1(chunk, test_fn, line);          // pushes I32(0 or 1)
    chunk.emit_op(Op::I32_EQZ, line);     // Bool(true) if 0, Bool(false) if 1
    chunk.emit_jump(Op::BR_IF_TRUE, line) // branch when not this type
}

/// AND two I32 type-test results, then branch to `skip_label` if AND=0.
fn and_skip_if_not(chunk: &mut Chunk, line: u32) -> usize {
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_jump(Op::BR_IF_TRUE, line)
}

// ── emit_dyn_to_bool ───────────────────────────────────────────────────

/// Truthy coercion — ECMA-262 §7.1.2 ToBoolean.
/// Stack: [v] → [i32: 0 or 1]
pub fn emit_dyn_to_bool(chunk: &mut Chunk, line: u32) {
    let slots = alloc_locals(chunk, 2);
    let v = slots;
    let f = slots + 1;

    let test_bool   = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool   = chunk.add_import("wasm:js-boolean", "cast");
    let test_num    = chunk.add_import("wasm:js-number",  "test");
    let to_f64      = chunk.add_import("wasm:js-number",  "toF64");
    let test_str    = chunk.add_import("wasm:js-string",  "test");
    let str_length  = chunk.add_import("wasm:js-string",  "length");
    let test_bigint = chunk.add_import("wasm:js-bigint",  "test");

    save(chunk, v, line);

    // null or undefined → false  (REF_IS_NULL returns Bool directly)
    load(chunk, v, line);
    chunk.emit_op(Op::REF_IS_NULL, line);       // Bool(true) = null/undef
    let null_false = chunk.emit_jump(Op::BR_IF_TRUE, line);

    // boolean: test returns I32 → cast returns I32 (0 or 1)
    let not_bool = test_and_skip_if_not(chunk, v, test_bool, line);
    load(chunk, v, line);
    call1(chunk, cast_bool, line);              // pushes I32(0 or 1)
    let done_bool = chunk.emit_jump(Op::BR, line);

    // number → f64 != 0.0 && !NaN
    chunk.patch_jump(not_bool);
    let not_num = test_and_skip_if_not(chunk, v, test_num, line);
    load(chunk, v, line);
    call1(chunk, to_f64, line);
    save(chunk, f, line);
    load(chunk, f, line);
    load(chunk, f, line);
    chunk.emit_op(Op::F64_NE, line);            // Bool(true) if NaN
    let nan_is_false = chunk.emit_jump(Op::BR_IF_TRUE, line);
    load(chunk, f, line);
    f64_const(chunk, 0.0, line);
    chunk.emit_op(Op::F64_NE, line);            // Bool: f != 0.0
    let done_num = chunk.emit_jump(Op::BR, line);

    // string → length > 0
    chunk.patch_jump(not_num);
    let not_str = test_and_skip_if_not(chunk, v, test_str, line);
    load(chunk, v, line);
    call1(chunk, str_length, line);             // I32
    i32_const(chunk, 0, line);
    chunk.emit_op(Op::I32_NE, line);            // Bool
    let done_str = chunk.emit_jump(Op::BR, line);

    // bigint → i64 != 0
    chunk.patch_jump(not_str);
    let not_bigint = test_and_skip_if_not(chunk, v, test_bigint, line);
    load(chunk, v, line);
    i64_const(chunk, 0, line);
    chunk.emit_op(Op::I64_NE, line);            // Bool
    let done_bigint = chunk.emit_jump(Op::BR, line);

    // anything else (object/symbol) → true
    chunk.patch_jump(not_bigint);
    i32_const(chunk, 1, line);
    let done_obj = chunk.emit_jump(Op::BR, line);

    // false paths
    chunk.patch_jump(null_false);
    chunk.patch_jump(nan_is_false);
    i32_const(chunk, 0, line);

    chunk.patch_jump(done_bool);
    chunk.patch_jump(done_num);
    chunk.patch_jump(done_str);
    chunk.patch_jump(done_bigint);
    chunk.patch_jump(done_obj);
}

// ── emit_dyn_not ──────────────────────────────────────────────────────

pub fn emit_dyn_not(chunk: &mut Chunk, line: u32) {
    emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
}

// ── emit_dyn_eq ───────────────────────────────────────────────────────

pub fn emit_dyn_eq(chunk: &mut Chunk, line: u32) {
    let slots = alloc_locals(chunk, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_num    = chunk.add_import("wasm:js-number",  "test");
    let to_f64      = chunk.add_import("wasm:js-number",  "toF64");
    let test_str    = chunk.add_import("wasm:js-string",  "test");
    let str_eq      = chunk.add_import("wasm:js-string",  "equals");
    let test_bool   = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool   = chunk.add_import("wasm:js-boolean", "cast");
    let test_bigint = chunk.add_import("wasm:js-bigint",  "test");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    // null/undefined: if a is nullish AND b is nullish → true; a nullish + b not → false
    load(chunk, a_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);               // Bool
    let a_not_nullish = chunk.emit_jump(Op::BR_IF_FALSE, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);               // Bool
    let both_nullish = chunk.emit_jump(Op::BR_IF_TRUE, line);
    let nullish_ne = chunk.emit_jump(Op::BR, line);     // a null, b not → false

    // both number?
    chunk.patch_jump(a_not_nullish);
    load(chunk, a_slot, line); call1(chunk, test_num, line);
    load(chunk, b_slot, line); call1(chunk, test_num, line);
    let not_num = and_skip_if_not(chunk, line);
    load(chunk, a_slot, line); call1(chunk, to_f64, line);
    load(chunk, b_slot, line); call1(chunk, to_f64, line);
    chunk.emit_op(Op::F64_EQ, line);
    let done_num = chunk.emit_jump(Op::BR, line);

    // both string?
    chunk.patch_jump(not_num);
    load(chunk, a_slot, line); call1(chunk, test_str, line);
    load(chunk, b_slot, line); call1(chunk, test_str, line);
    let not_str = and_skip_if_not(chunk, line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    call2(chunk, str_eq, line);
    let done_str = chunk.emit_jump(Op::BR, line);

    // both boolean?
    chunk.patch_jump(not_str);
    load(chunk, a_slot, line); call1(chunk, test_bool, line);
    load(chunk, b_slot, line); call1(chunk, test_bool, line);
    let not_bool = and_skip_if_not(chunk, line);
    load(chunk, a_slot, line); call1(chunk, cast_bool, line);
    load(chunk, b_slot, line); call1(chunk, cast_bool, line);
    chunk.emit_op(Op::I32_EQ, line);
    let done_bool = chunk.emit_jump(Op::BR, line);

    // both bigint?
    chunk.patch_jump(not_bool);
    load(chunk, a_slot, line); call1(chunk, test_bigint, line);
    load(chunk, b_slot, line); call1(chunk, test_bigint, line);
    let not_bigint = and_skip_if_not(chunk, line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::I64_EQ, line);
    let done_bigint = chunk.emit_jump(Op::BR, line);

    // object / cross-type → ref.eq
    chunk.patch_jump(not_bigint);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::REF_EQ, line);
    let done_ref = chunk.emit_jump(Op::BR, line);

    // true
    chunk.patch_jump(both_nullish);
    i32_const(chunk, 1, line);
    let done_true = chunk.emit_jump(Op::BR, line);

    // false (nullish + non-nullish)
    chunk.patch_jump(nullish_ne);
    i32_const(chunk, 0, line);

    chunk.patch_jump(done_num);
    chunk.patch_jump(done_str);
    chunk.patch_jump(done_bool);
    chunk.patch_jump(done_bigint);
    chunk.patch_jump(done_ref);
    chunk.patch_jump(done_true);
}

pub fn emit_dyn_ne(chunk: &mut Chunk, line: u32) {
    emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
}

// ── comparison helpers ────────────────────────────────────────────────

enum CmpOp { Lt, Gt, Le, Ge }

fn f64_cmp_op(op: &CmpOp) -> Op {
    match op { CmpOp::Lt => Op::F64_LT, CmpOp::Gt => Op::F64_GT,
               CmpOp::Le => Op::F64_LE, CmpOp::Ge => Op::F64_GE }
}

fn i32_cmp_op(op: &CmpOp) -> Op {
    match op { CmpOp::Lt => Op::I32_LT_S, CmpOp::Gt => Op::I32_GT_S,
               CmpOp::Le => Op::I32_LE_S, CmpOp::Ge => Op::I32_GE_S }
}

fn i64_cmp_op(op: &CmpOp) -> Op {
    match op { CmpOp::Lt => Op::I64_LT_S, CmpOp::Gt => Op::I64_GT_S,
               CmpOp::Le => Op::I64_LE_S, CmpOp::Ge => Op::I64_GE_S }
}

fn emit_dyn_cmp(chunk: &mut Chunk, line: u32, op: CmpOp) {
    let slots = alloc_locals(chunk, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_num    = chunk.add_import("wasm:js-number", "test");
    let to_f64      = chunk.add_import("wasm:js-number", "toF64");
    let test_str    = chunk.add_import("wasm:js-string", "test");
    let str_compare = chunk.add_import("wasm:js-string", "compare");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    // both number?
    load(chunk, a_slot, line); call1(chunk, test_num, line);
    load(chunk, b_slot, line); call1(chunk, test_num, line);
    let not_num = and_skip_if_not(chunk, line);
    load(chunk, a_slot, line); call1(chunk, to_f64, line);
    load(chunk, b_slot, line); call1(chunk, to_f64, line);
    chunk.emit_op(f64_cmp_op(&op), line);
    let done_num = chunk.emit_jump(Op::BR, line);

    // both string? → compare returns I32 (-1/0/1)
    chunk.patch_jump(not_num);
    load(chunk, a_slot, line); call1(chunk, test_str, line);
    load(chunk, b_slot, line); call1(chunk, test_str, line);
    let not_str = and_skip_if_not(chunk, line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    call2(chunk, str_compare, line);        // I32 (-1/0/1)
    i32_const(chunk, 0, line);
    chunk.emit_op(i32_cmp_op(&op), line);   // Bool
    let done_str = chunk.emit_jump(Op::BR, line);

    // both bigint?
    chunk.patch_jump(not_str);
    load(chunk, a_slot, line); call1(chunk, test_bigint, line);
    load(chunk, b_slot, line); call1(chunk, test_bigint, line);
    let not_bigint = and_skip_if_not(chunk, line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(i64_cmp_op(&op), line);   // Bool
    let done_bigint = chunk.emit_jump(Op::BR, line);

    // fallback: coerce both to f64
    chunk.patch_jump(not_bigint);
    load(chunk, a_slot, line); call1(chunk, to_f64, line);
    load(chunk, b_slot, line); call1(chunk, to_f64, line);
    chunk.emit_op(f64_cmp_op(&op), line);

    chunk.patch_jump(done_num);
    chunk.patch_jump(done_str);
    chunk.patch_jump(done_bigint);
}

pub fn emit_dyn_lt(chunk: &mut Chunk, line: u32) { emit_dyn_cmp(chunk, line, CmpOp::Lt); }
pub fn emit_dyn_gt(chunk: &mut Chunk, line: u32) { emit_dyn_cmp(chunk, line, CmpOp::Gt); }
pub fn emit_dyn_le(chunk: &mut Chunk, line: u32) { emit_dyn_cmp(chunk, line, CmpOp::Le); }
pub fn emit_dyn_ge(chunk: &mut Chunk, line: u32) { emit_dyn_cmp(chunk, line, CmpOp::Ge); }

// ── emit_dyn_add ──────────────────────────────────────────────────────

pub fn emit_dyn_add(chunk: &mut Chunk, line: u32) {
    let slots = alloc_locals(chunk, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_str     = chunk.add_import("wasm:js-string", "test");
    let str_cast     = chunk.add_import("wasm:js-string", "cast");
    let str_concat   = chunk.add_import("wasm:js-string", "concat");
    let str_from_f64 = chunk.add_import("wasm:js-string", "fromF64");
    let test_bigint  = chunk.add_import("wasm:js-bigint", "test");
    let test_num     = chunk.add_import("wasm:js-number", "test");
    let to_f64       = chunk.add_import("wasm:js-number", "toF64");
    let from_f64     = chunk.add_import("wasm:js-number", "fromF64");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    // either is a string → coerce both to string, then concat
    load(chunk, a_slot, line); call1(chunk, test_str, line);
    load(chunk, b_slot, line); call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op(Op::I32_EQZ, line);
    let not_str = chunk.emit_jump(Op::BR_IF_TRUE, line);

    // coerce a to string
    load(chunk, a_slot, line); call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_EQZ, line);
    let a_already_str = chunk.emit_jump(Op::BR_IF_FALSE, line);
    load(chunk, a_slot, line); call1(chunk, to_f64, line); call1(chunk, str_from_f64, line);
    let a_coerced = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(a_already_str);
    load(chunk, a_slot, line); call1(chunk, str_cast, line);
    chunk.patch_jump(a_coerced);

    // coerce b to string
    load(chunk, b_slot, line); call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_EQZ, line);
    let b_already_str = chunk.emit_jump(Op::BR_IF_FALSE, line);
    load(chunk, b_slot, line); call1(chunk, to_f64, line); call1(chunk, str_from_f64, line);
    let b_coerced = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(b_already_str);
    load(chunk, b_slot, line); call1(chunk, str_cast, line);
    chunk.patch_jump(b_coerced);

    call2(chunk, str_concat, line);
    let done_str = chunk.emit_jump(Op::BR, line);

    // both bigint → i64.add
    chunk.patch_jump(not_str);
    load(chunk, a_slot, line); call1(chunk, test_bigint, line);
    load(chunk, b_slot, line); call1(chunk, test_bigint, line);
    let not_bigint = and_skip_if_not(chunk, line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::I64_ADD, line);
    let done_bigint = chunk.emit_jump(Op::BR, line);

    // number + number → f64.add
    chunk.patch_jump(not_bigint);
    load(chunk, a_slot, line); call1(chunk, to_f64, line);
    load(chunk, b_slot, line); call1(chunk, to_f64, line);
    chunk.emit_op(Op::F64_ADD, line);
    call1(chunk, from_f64, line);

    chunk.patch_jump(done_str);
    chunk.patch_jump(done_bigint);
}

// ── emit_dyn_neg ──────────────────────────────────────────────────────

pub fn emit_dyn_neg(chunk: &mut Chunk, line: u32) {
    let v = alloc_locals(chunk, 1);

    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let to_f64      = chunk.add_import("wasm:js-number", "toF64");
    let from_f64    = chunk.add_import("wasm:js-number", "fromF64");

    save(chunk, v, line);

    // bigint → i64 negation (0 - v)
    let not_bigint = test_and_skip_if_not(chunk, v, test_bigint, line);
    i64_const(chunk, 0, line);
    load(chunk, v, line);
    chunk.emit_op(Op::I64_SUB, line);
    let done_bigint = chunk.emit_jump(Op::BR, line);

    // number → f64 negation
    chunk.patch_jump(not_bigint);
    load(chunk, v, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(Op::F64_NEG, line);
    call1(chunk, from_f64, line);

    chunk.patch_jump(done_bigint);
}
