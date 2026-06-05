//! Dynamic-dispatch opcode emitters — spec-compliant replacements for DYN_*.
//!
//! Each function emits the WASM-standard type-dispatch sequence using only:
//!   - Standard WASM opcodes (Op::IF / Op::ELSE / Op::END for control flow)
//!   - `wasm:js-*` host imports (js-string-builtins + js-primitive-builtins proposals)
//!
//! Control flow uses actual WASM structured control flow:
//!   `Op::IF` (0x04) / `Op::ELSE` (0x05) / `Op::END` (0x0B)
//! No flat-offset BR_IF_FALSE / BR_IF_TRUE / BR_IF_NULL custom opcodes.

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

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

/// Convert the i32 (or Bool) on top of stack to a canonical JS Bool.
/// Stack: [i32 or Bool] → [Bool(true) if nonzero/true, Bool(false) otherwise]
fn i32_to_bool(chunk: &mut Chunk, line: u32) {
    chunk.emit_if_value(line);
    chunk.emit_op(Op::TRUE, line);
    chunk.emit_else(line);
    chunk.emit_op(Op::FALSE, line);
    chunk.emit_end(line);
}

/// Convert an `i32` (0 or 1) on the stack to a JS `Value::Bool`.
///
/// Used wherever a WASM-level comparison opcode (`REF_IS_*`, `emit_dyn_eq`,
/// etc.) produces `i32` but the ECMA-262 runtime expects `boolean`.
/// Branch conditions (`if`/`br_if`/`emit_loop_cond`) accept both `i32` and
/// `Bool`, so this wrapper is only needed in *value* positions.
pub fn emit_i32_to_bool(chunk: &mut Chunk, line: u32) {
    i32_to_bool(chunk, line);
}

fn save(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
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

// ── emit_dyn_to_bool ───────────────────────────────────────────────────

/// Truthy coercion — ECMA-262 §7.1.2 ToBoolean.
/// Stack: [v] → [i32: 0 or 1]
/// Uses WASM structured control flow: Op::IF / Op::ELSE / Op::END.
pub fn emit_dyn_to_bool(chunk: &mut Chunk, line: u32) {
    let slots = alloc_locals(chunk, 2);
    let v = slots;
    let f = slots + 1;

    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let test_num = chunk.add_import("wasm:js-number", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_length = chunk.add_import("wasm:js-string", "length");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");

    save(chunk, v, line);

    // null / undefined → false
    load(chunk, v, line);
    chunk.emit_op(Op::REF_IS_NULL, line); // i32: 1 if null
    chunk.emit_if(line);
    i32_const(chunk, 0, line);
    chunk.emit_else(line);

    // boolean? — cast_bool returns i32 (1=true, 0=false) for a known Bool value
    load(chunk, v, line);
    call1(chunk, test_bool, line); // i32: 1 if bool
    chunk.emit_if(line);
    load(chunk, v, line);
    call1(chunk, cast_bool, line); // Bool → i32
    chunk.emit_else(line);

    // number?
    load(chunk, v, line);
    call1(chunk, test_num, line); // i32: 1 if number
    chunk.emit_if(line);
    load(chunk, v, line);
    call1(chunk, to_f64, line); // f64
    save(chunk, f, line);
    load(chunk, f, line);
    load(chunk, f, line);
    chunk.emit_op(Op::F64_NE, line); // i32: 1 if NaN (NaN != NaN)
    chunk.emit_if(line); // NaN → false
    i32_const(chunk, 0, line);
    chunk.emit_else(line);
    load(chunk, f, line);
    f64_const(chunk, 0.0, line);
    chunk.emit_op(Op::F64_NE, line); // i32: 1 if nonzero
    chunk.emit_end(line);

    chunk.emit_else(line);

    // string?
    load(chunk, v, line);
    call1(chunk, test_str, line); // i32: 1 if string
    chunk.emit_if(line);
    load(chunk, v, line);
    call1(chunk, str_length, line); // i32 length
    i32_const(chunk, 0, line);
    chunk.emit_op(Op::I32_NE, line); // i32: 1 if nonempty
    chunk.emit_else(line);

    // bigint?
    load(chunk, v, line);
    call1(chunk, test_bigint, line); // i32: 1 if bigint
    chunk.emit_if(line);
    load(chunk, v, line);
    i64_const(chunk, 0, line);
    chunk.emit_op(Op::I64_NE, line); // i32: 1 if nonzero
    chunk.emit_else(line);
    i32_const(chunk, 1, line); // object / symbol → truthy
    chunk.emit_end(line);

    chunk.emit_end(line); // end string
    chunk.emit_end(line); // end number
    chunk.emit_end(line); // end boolean
    chunk.emit_end(line); // end null
}

// ── emit_dyn_not ──────────────────────────────────────────────────────

pub fn emit_dyn_not(chunk: &mut Chunk, line: u32) {
    emit_dyn_to_bool(chunk, line); // any → i32
    chunk.emit_op(Op::I32_EQZ, line); // negate: i32
    // Result is i32 — WASM-compliant
}

// ── emit_dyn_eq ───────────────────────────────────────────────────────

pub fn emit_dyn_eq(chunk: &mut Chunk, line: u32) {
    let slots = alloc_locals(chunk, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_num = chunk.add_import("wasm:js-number", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_eq = chunk.add_import("wasm:js-string", "equals");
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    // a is null/undefined?
    load(chunk, a_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line); // i32
    chunk.emit_if(line);
    // a is null → true iff b is also null
    load(chunk, b_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line); // i32
    chunk.emit_if(line);
    i32_const(chunk, 1, line); // both null → equal
    chunk.emit_else(line);
    i32_const(chunk, 0, line); // a null, b not → not equal
    chunk.emit_end(line);

    chunk.emit_else(line);

    // both number?
    load(chunk, a_slot, line);
    call1(chunk, test_num, line);
    load(chunk, b_slot, line);
    call1(chunk, test_num, line);
    chunk.emit_op(Op::I32_AND, line); // i32: 1 if both numbers
    chunk.emit_if(line);
    load(chunk, a_slot, line);
    call1(chunk, to_f64, line);
    load(chunk, b_slot, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_else(line);

    // both string?
    load(chunk, a_slot, line);
    call1(chunk, test_str, line);
    load(chunk, b_slot, line);
    call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_AND, line); // i32: 1 if both strings
    chunk.emit_if(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    call2(chunk, str_eq, line);
    chunk.emit_else(line);

    // both boolean?
    load(chunk, a_slot, line);
    call1(chunk, test_bool, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bool, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    // cast_bool(v) → i32 (1=true, 0=false) for known Bool values
    load(chunk, a_slot, line);
    call1(chunk, cast_bool, line);
    load(chunk, b_slot, line);
    call1(chunk, cast_bool, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_else(line);

    // both bigint?
    load(chunk, a_slot, line);
    call1(chunk, test_bigint, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bigint, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::I64_EQ, line);
    chunk.emit_else(line);
    // object / cross-type → reference equality
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::REF_EQ, line);
    chunk.emit_end(line); // bigint
    chunk.emit_end(line); // boolean
    chunk.emit_end(line); // string
    chunk.emit_end(line); // number
    chunk.emit_end(line); // null
    // Result is i32 (0 or 1) — WASM-compliant for IF conditions
}

pub fn emit_dyn_ne(chunk: &mut Chunk, line: u32) {
    emit_dyn_eq(chunk, line); // i32
    chunk.emit_op(Op::I32_EQZ, line); // negate: 1 if not-equal
}

// ── comparison helpers ────────────────────────────────────────────────

enum CmpOp {
    Lt,
    Gt,
    Le,
    Ge,
}

fn f64_cmp_op(op: &CmpOp) -> Op {
    match op {
        CmpOp::Lt => Op::F64_LT,
        CmpOp::Gt => Op::F64_GT,
        CmpOp::Le => Op::F64_LE,
        CmpOp::Ge => Op::F64_GE,
    }
}

fn i32_cmp_op(op: &CmpOp) -> Op {
    match op {
        CmpOp::Lt => Op::I32_LT_S,
        CmpOp::Gt => Op::I32_GT_S,
        CmpOp::Le => Op::I32_LE_S,
        CmpOp::Ge => Op::I32_GE_S,
    }
}

fn i64_cmp_op(op: &CmpOp) -> Op {
    match op {
        CmpOp::Lt => Op::I64_LT_S,
        CmpOp::Gt => Op::I64_GT_S,
        CmpOp::Le => Op::I64_LE_S,
        CmpOp::Ge => Op::I64_GE_S,
    }
}

fn emit_dyn_cmp(chunk: &mut Chunk, line: u32, op: CmpOp) {
    let slots = alloc_locals(chunk, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_num = chunk.add_import("wasm:js-number", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_compare = chunk.add_import("wasm:js-string", "compare");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    // both number?
    load(chunk, a_slot, line);
    call1(chunk, test_num, line);
    load(chunk, b_slot, line);
    call1(chunk, test_num, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    load(chunk, a_slot, line);
    call1(chunk, to_f64, line);
    load(chunk, b_slot, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(f64_cmp_op(&op), line);
    chunk.emit_else(line);

    // both string?
    load(chunk, a_slot, line);
    call1(chunk, test_str, line);
    load(chunk, b_slot, line);
    call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    call2(chunk, str_compare, line); // i32 (-1/0/1)
    i32_const(chunk, 0, line);
    chunk.emit_op(i32_cmp_op(&op), line);
    chunk.emit_else(line);

    // both bigint?
    load(chunk, a_slot, line);
    call1(chunk, test_bigint, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bigint, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(i64_cmp_op(&op), line);
    chunk.emit_else(line);
    // fallback: coerce both to f64
    load(chunk, a_slot, line);
    call1(chunk, to_f64, line);
    load(chunk, b_slot, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(f64_cmp_op(&op), line);
    chunk.emit_end(line); // bigint
    chunk.emit_end(line); // string
    chunk.emit_end(line); // number
    // Result is i32 (0 or 1) — WASM-compliant for IF/BR_IF conditions
}

pub fn emit_dyn_lt(chunk: &mut Chunk, line: u32) {
    emit_dyn_cmp(chunk, line, CmpOp::Lt);
}
pub fn emit_dyn_gt(chunk: &mut Chunk, line: u32) {
    emit_dyn_cmp(chunk, line, CmpOp::Gt);
}
pub fn emit_dyn_le(chunk: &mut Chunk, line: u32) {
    emit_dyn_cmp(chunk, line, CmpOp::Le);
}
pub fn emit_dyn_ge(chunk: &mut Chunk, line: u32) {
    emit_dyn_cmp(chunk, line, CmpOp::Ge);
}

// ── emit_dyn_add ──────────────────────────────────────────────────────

pub fn emit_dyn_add(chunk: &mut Chunk, line: u32) {
    let slots = alloc_locals(chunk, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_cast = chunk.add_import("wasm:js-string", "cast");
    let str_concat = chunk.add_import("wasm:js-string", "concat");
    let str_from_f64 = chunk.add_import("wasm:js-string", "fromF64");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let test_num = chunk.add_import("wasm:js-number", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let from_f64 = chunk.add_import("wasm:js-number", "fromF64");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    // either is a string → string concatenation
    load(chunk, a_slot, line);
    call1(chunk, test_str, line);
    load(chunk, b_slot, line);
    call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_OR, line); // i32: 1 if either is string
    chunk.emit_if(line);
    // coerce a to string
    load(chunk, a_slot, line);
    call1(chunk, test_str, line); // i32: 1 if a is already string
    chunk.emit_if(line);
    load(chunk, a_slot, line);
    call1(chunk, str_cast, line);
    chunk.emit_else(line);
    load(chunk, a_slot, line);
    call1(chunk, to_f64, line);
    call1(chunk, str_from_f64, line);
    chunk.emit_end(line);
    // coerce b to string
    load(chunk, b_slot, line);
    call1(chunk, test_str, line);
    chunk.emit_if(line);
    load(chunk, b_slot, line);
    call1(chunk, str_cast, line);
    chunk.emit_else(line);
    load(chunk, b_slot, line);
    call1(chunk, to_f64, line);
    call1(chunk, str_from_f64, line);
    chunk.emit_end(line);
    call2(chunk, str_concat, line);

    chunk.emit_else(line);
    // both bigint → i64.add
    load(chunk, a_slot, line);
    call1(chunk, test_bigint, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bigint, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    chunk.emit_op(Op::I64_ADD, line);
    chunk.emit_else(line);
    // number + number (or coerce) → f64.add
    load(chunk, a_slot, line);
    call1(chunk, to_f64, line);
    load(chunk, b_slot, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(Op::F64_ADD, line);
    call1(chunk, from_f64, line);
    chunk.emit_end(line); // bigint
    chunk.emit_end(line); // string
}

// ── emit_dyn_neg ──────────────────────────────────────────────────────

pub fn emit_dyn_neg(chunk: &mut Chunk, line: u32) {
    let v = alloc_locals(chunk, 1);

    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let from_f64 = chunk.add_import("wasm:js-number", "fromF64");

    save(chunk, v, line);

    // bigint → i64 negation
    load(chunk, v, line);
    call1(chunk, test_bigint, line); // i32: 1 if bigint
    chunk.emit_if(line);
    i64_const(chunk, 0, line);
    load(chunk, v, line);
    chunk.emit_op(Op::I64_SUB, line);
    chunk.emit_else(line);
    // number → f64 negation
    load(chunk, v, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(Op::F64_NEG, line);
    call1(chunk, from_f64, line);
    chunk.emit_end(line);
}

// ── _into variants ────────────────────────────────────────────────────────
//
// These register imports on `imports` (the module-level chunk, chunks[0]) and
// emit bytecode into `code` (the function chunk being built).  This separation
// is REQUIRED for WASM compliance: a WASM module has ONE import section.
//
// Without this split, every emit_dyn_* call on a code chunk registers imports
// locally on that chunk (via chunk.add_import).  When emit_import_call_into
// also emits CALL_IMPORT using the module-level index, the two index spaces
// collide: CALL_IMPORT K in the code chunk resolves to the code chunk's local
// import K (e.g. wasm:js-boolean.test) instead of the module-level import K
// (e.g. ecma:array.new).  Use the _into variants in any stdlib builder that
// has a separate `imports: &mut Chunk` parameter.

pub fn emit_dyn_to_bool_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let slots = alloc_locals(code, 2);
    let v = slots;
    let f = slots + 1;

    let test_bool = imports.add_import("wasm:js-boolean", "test");
    let cast_bool = imports.add_import("wasm:js-boolean", "cast");
    let test_num = imports.add_import("wasm:js-number", "test");
    let to_f64 = imports.add_import("wasm:js-number", "toF64");
    let test_str = imports.add_import("wasm:js-string", "test");
    let str_length = imports.add_import("wasm:js-string", "length");
    let test_bigint = imports.add_import("wasm:js-bigint", "test");

    save(code, v, line);

    load(code, v, line);
    code.emit_op(Op::REF_IS_NULL, line);
    code.emit_if(line);
    i32_const(code, 0, line);
    code.emit_else(line);

    load(code, v, line);
    call1(code, test_bool, line);
    code.emit_if(line);
    load(code, v, line);
    call1(code, cast_bool, line); // Bool → i32
    code.emit_else(line);

    load(code, v, line);
    call1(code, test_num, line);
    code.emit_if(line);
    load(code, v, line);
    call1(code, to_f64, line);
    save(code, f, line);
    load(code, f, line);
    load(code, f, line);
    code.emit_op(Op::F64_NE, line);
    code.emit_if(line);
    i32_const(code, 0, line);
    code.emit_else(line);
    load(code, f, line);
    f64_const(code, 0.0, line);
    code.emit_op(Op::F64_NE, line);
    code.emit_end(line);
    code.emit_else(line);

    load(code, v, line);
    call1(code, test_str, line);
    code.emit_if(line);
    load(code, v, line);
    call1(code, str_length, line);
    i32_const(code, 0, line);
    code.emit_op(Op::I32_NE, line);
    code.emit_else(line);

    load(code, v, line);
    call1(code, test_bigint, line);
    code.emit_if(line);
    load(code, v, line);
    i64_const(code, 0, line);
    code.emit_op(Op::I64_NE, line);
    code.emit_else(line);
    i32_const(code, 1, line);
    code.emit_end(line);

    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
}

pub fn emit_dyn_not_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_to_bool_into(imports, code, line);
    code.emit_op(Op::I32_EQZ, line); // i32
}

pub fn emit_dyn_eq_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let slots = alloc_locals(code, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_num = imports.add_import("wasm:js-number", "test");
    let to_f64 = imports.add_import("wasm:js-number", "toF64");
    let test_str = imports.add_import("wasm:js-string", "test");
    let str_eq = imports.add_import("wasm:js-string", "equals");
    let test_bool = imports.add_import("wasm:js-boolean", "test");
    let cast_bool = imports.add_import("wasm:js-boolean", "cast");
    let test_bigint = imports.add_import("wasm:js-bigint", "test");

    save(code, b_slot, line);
    save(code, a_slot, line);

    load(code, a_slot, line);
    code.emit_op(Op::REF_IS_NULL, line);
    code.emit_if(line);
    load(code, b_slot, line);
    code.emit_op(Op::REF_IS_NULL, line);
    code.emit_if(line);
    i32_const(code, 1, line);
    code.emit_else(line);
    i32_const(code, 0, line);
    code.emit_end(line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_num, line);
    load(code, b_slot, line);
    call1(code, test_num, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if(line);
    load(code, a_slot, line);
    call1(code, to_f64, line);
    load(code, b_slot, line);
    call1(code, to_f64, line);
    code.emit_op(Op::F64_EQ, line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_str, line);
    load(code, b_slot, line);
    call1(code, test_str, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    call2(code, str_eq, line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_bool, line);
    load(code, b_slot, line);
    call1(code, test_bool, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if(line);
    load(code, a_slot, line);
    call1(code, cast_bool, line); // Bool → i32
    load(code, b_slot, line);
    call1(code, cast_bool, line); // Bool → i32
    code.emit_op(Op::I32_EQ, line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_bigint, line);
    load(code, b_slot, line);
    call1(code, test_bigint, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    code.emit_op(Op::I64_EQ, line);
    code.emit_else(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    code.emit_op(Op::REF_EQ, line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    // Result is i32 — WASM-compliant
}

pub fn emit_dyn_ne_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_eq_into(imports, code, line); // i32
    code.emit_op(Op::I32_EQZ, line); // negate: 1 if not-equal, i32
}

fn emit_dyn_cmp_into(imports: &mut Chunk, code: &mut Chunk, line: u32, op: CmpOp) {
    let slots = alloc_locals(code, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_num = imports.add_import("wasm:js-number", "test");
    let to_f64 = imports.add_import("wasm:js-number", "toF64");
    let test_str = imports.add_import("wasm:js-string", "test");
    let str_compare = imports.add_import("wasm:js-string", "compare");
    let test_bigint = imports.add_import("wasm:js-bigint", "test");

    save(code, b_slot, line);
    save(code, a_slot, line);

    load(code, a_slot, line);
    call1(code, test_num, line);
    load(code, b_slot, line);
    call1(code, test_num, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if(line);
    load(code, a_slot, line);
    call1(code, to_f64, line);
    load(code, b_slot, line);
    call1(code, to_f64, line);
    code.emit_op(f64_cmp_op(&op), line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_str, line);
    load(code, b_slot, line);
    call1(code, test_str, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    call2(code, str_compare, line);
    i32_const(code, 0, line);
    code.emit_op(i32_cmp_op(&op), line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_bigint, line);
    load(code, b_slot, line);
    call1(code, test_bigint, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    code.emit_op(i64_cmp_op(&op), line);
    code.emit_else(line);
    load(code, a_slot, line);
    call1(code, to_f64, line);
    load(code, b_slot, line);
    call1(code, to_f64, line);
    code.emit_op(f64_cmp_op(&op), line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    // Result is i32 — WASM-compliant
}

pub fn emit_dyn_lt_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_cmp_into(imports, code, line, CmpOp::Lt);
}
pub fn emit_dyn_gt_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_cmp_into(imports, code, line, CmpOp::Gt);
}
pub fn emit_dyn_le_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_cmp_into(imports, code, line, CmpOp::Le);
}
pub fn emit_dyn_ge_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    emit_dyn_cmp_into(imports, code, line, CmpOp::Ge);
}

pub fn emit_dyn_add_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let slots = alloc_locals(code, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_str = imports.add_import("wasm:js-string", "test");
    let str_cast = imports.add_import("wasm:js-string", "cast");
    let str_concat = imports.add_import("wasm:js-string", "concat");
    let str_from_f64 = imports.add_import("wasm:js-string", "fromF64");
    let test_bigint = imports.add_import("wasm:js-bigint", "test");
    let test_num = imports.add_import("wasm:js-number", "test");
    let to_f64 = imports.add_import("wasm:js-number", "toF64");
    let from_f64 = imports.add_import("wasm:js-number", "fromF64");

    save(code, b_slot, line);
    save(code, a_slot, line);

    load(code, a_slot, line);
    call1(code, test_str, line);
    load(code, b_slot, line);
    call1(code, test_str, line);
    code.emit_op(Op::I32_OR, line);
    code.emit_if(line);
    load(code, a_slot, line);
    call1(code, test_str, line);
    code.emit_if(line);
    load(code, a_slot, line);
    call1(code, str_cast, line);
    code.emit_else(line);
    load(code, a_slot, line);
    call1(code, to_f64, line);
    call1(code, str_from_f64, line);
    code.emit_end(line);
    load(code, b_slot, line);
    call1(code, test_str, line);
    code.emit_if(line);
    load(code, b_slot, line);
    call1(code, str_cast, line);
    code.emit_else(line);
    load(code, b_slot, line);
    call1(code, to_f64, line);
    call1(code, str_from_f64, line);
    code.emit_end(line);
    call2(code, str_concat, line);

    code.emit_else(line);
    load(code, a_slot, line);
    call1(code, test_bigint, line);
    load(code, b_slot, line);
    call1(code, test_bigint, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    code.emit_op(Op::I64_ADD, line);
    code.emit_else(line);
    load(code, a_slot, line);
    call1(code, to_f64, line);
    load(code, b_slot, line);
    call1(code, to_f64, line);
    code.emit_op(Op::F64_ADD, line);
    call1(code, from_f64, line);
    code.emit_end(line);
    code.emit_end(line);
}

pub fn emit_dyn_neg_into(imports: &mut Chunk, code: &mut Chunk, line: u32) {
    let v = alloc_locals(code, 1);

    let test_bigint = imports.add_import("wasm:js-bigint", "test");
    let to_f64 = imports.add_import("wasm:js-number", "toF64");
    let from_f64 = imports.add_import("wasm:js-number", "fromF64");

    save(code, v, line);

    load(code, v, line);
    call1(code, test_bigint, line);
    code.emit_if(line);
    i64_const(code, 0, line);
    load(code, v, line);
    code.emit_op(Op::I64_SUB, line);
    code.emit_else(line);
    load(code, v, line);
    call1(code, to_f64, line);
    code.emit_op(Op::F64_NEG, line);
    call1(code, from_f64, line);
    code.emit_end(line);
}
