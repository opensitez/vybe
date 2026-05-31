//! Dynamic-dispatch opcode emitters — spec-compliant replacements for DYN_*.
//!
//! Each function emits the WASM-standard type-dispatch sequence using only:
//!   - Standard WASM opcodes (Op::IF / Op::ELSE / Op::END for control flow)
//!   - `wasm:js-*` host imports (js-string-builtins + js-primitive-builtins proposals)
//!
//! Control flow uses actual WASM structured control flow:
//!   `Op::IF` (0x04) / `Op::ELSE` (0x05) / `Op::END` (0x0B)
//! No flat-offset BR_IF_FALSE / BR_IF_TRUE / BR_IF_NULL custom opcodes.

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

    let test_bool   = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool   = chunk.add_import("wasm:js-boolean", "cast");
    let test_num    = chunk.add_import("wasm:js-number",  "test");
    let to_f64      = chunk.add_import("wasm:js-number",  "toF64");
    let test_str    = chunk.add_import("wasm:js-string",  "test");
    let str_length  = chunk.add_import("wasm:js-string",  "length");
    let test_bigint = chunk.add_import("wasm:js-bigint",  "test");

    save(chunk, v, line);

    // null / undefined → false
    load(chunk, v, line);
    chunk.emit_op(Op::REF_IS_NULL, line);   // i32: 1 if null
    chunk.emit_if(line);
      i32_const(chunk, 0, line);
    chunk.emit_else(line);

      // boolean?
      load(chunk, v, line);
      call1(chunk, test_bool, line);        // i32: 1 if bool
      chunk.emit_if(line);
        load(chunk, v, line);
        call1(chunk, cast_bool, line);      // i32 (0 or 1)
      chunk.emit_else(line);

        // number?
        load(chunk, v, line);
        call1(chunk, test_num, line);       // i32: 1 if number
        chunk.emit_if(line);
          load(chunk, v, line);
          call1(chunk, to_f64, line);       // f64
          save(chunk, f, line);
          load(chunk, f, line);
          load(chunk, f, line);
          chunk.emit_op(Op::F64_NE, line);  // i32: 1 if NaN (NaN != NaN)
          chunk.emit_if(line);              // NaN → false
            i32_const(chunk, 0, line);
          chunk.emit_else(line);
            load(chunk, f, line);
            f64_const(chunk, 0.0, line);
            chunk.emit_op(Op::F64_NE, line); // i32: 1 if nonzero
          chunk.emit_end(line);

        chunk.emit_else(line);

          // string?
          load(chunk, v, line);
          call1(chunk, test_str, line);     // i32: 1 if string
          chunk.emit_if(line);
            load(chunk, v, line);
            call1(chunk, str_length, line); // i32 length
            i32_const(chunk, 0, line);
            chunk.emit_op(Op::I32_NE, line);// i32: 1 if nonempty
          chunk.emit_else(line);

            // bigint?
            load(chunk, v, line);
            call1(chunk, test_bigint, line);// i32: 1 if bigint
            chunk.emit_if(line);
              load(chunk, v, line);
              i64_const(chunk, 0, line);
              chunk.emit_op(Op::I64_NE, line); // i32: 1 if nonzero
            chunk.emit_else(line);
              i32_const(chunk, 1, line);    // object / symbol → truthy
            chunk.emit_end(line);

          chunk.emit_end(line); // end string
        chunk.emit_end(line);   // end number
      chunk.emit_end(line);     // end boolean
    chunk.emit_end(line);       // end null
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

    // a is null/undefined?
    load(chunk, a_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);   // i32
    chunk.emit_if(line);
      // a is null → true iff b is also null
      load(chunk, b_slot, line);
      chunk.emit_op(Op::REF_IS_NULL, line); // i32
      chunk.emit_if(line);
        i32_const(chunk, 1, line);          // both null → equal
      chunk.emit_else(line);
        i32_const(chunk, 0, line);          // a null, b not → not equal
      chunk.emit_end(line);

    chunk.emit_else(line);

      // both number?
      load(chunk, a_slot, line); call1(chunk, test_num, line);
      load(chunk, b_slot, line); call1(chunk, test_num, line);
      chunk.emit_op(Op::I32_AND, line);     // i32: 1 if both numbers
      chunk.emit_if(line);
        load(chunk, a_slot, line); call1(chunk, to_f64, line);
        load(chunk, b_slot, line); call1(chunk, to_f64, line);
        chunk.emit_op(Op::F64_EQ, line);
      chunk.emit_else(line);

        // both string?
        load(chunk, a_slot, line); call1(chunk, test_str, line);
        load(chunk, b_slot, line); call1(chunk, test_str, line);
        chunk.emit_op(Op::I32_AND, line);   // i32: 1 if both strings
        chunk.emit_if(line);
          load(chunk, a_slot, line);
          load(chunk, b_slot, line);
          call2(chunk, str_eq, line);
        chunk.emit_else(line);

          // both boolean?
          load(chunk, a_slot, line); call1(chunk, test_bool, line);
          load(chunk, b_slot, line); call1(chunk, test_bool, line);
          chunk.emit_op(Op::I32_AND, line);
          chunk.emit_if(line);
            load(chunk, a_slot, line); call1(chunk, cast_bool, line);
            load(chunk, b_slot, line); call1(chunk, cast_bool, line);
            chunk.emit_op(Op::I32_EQ, line);
          chunk.emit_else(line);

            // both bigint?
            load(chunk, a_slot, line); call1(chunk, test_bigint, line);
            load(chunk, b_slot, line); call1(chunk, test_bigint, line);
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
          chunk.emit_end(line);   // boolean
        chunk.emit_end(line);     // string
      chunk.emit_end(line);       // number
    chunk.emit_end(line);         // null
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
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
      load(chunk, a_slot, line); call1(chunk, to_f64, line);
      load(chunk, b_slot, line); call1(chunk, to_f64, line);
      chunk.emit_op(f64_cmp_op(&op), line);
    chunk.emit_else(line);

      // both string?
      load(chunk, a_slot, line); call1(chunk, test_str, line);
      load(chunk, b_slot, line); call1(chunk, test_str, line);
      chunk.emit_op(Op::I32_AND, line);
      chunk.emit_if(line);
        load(chunk, a_slot, line);
        load(chunk, b_slot, line);
        call2(chunk, str_compare, line);    // i32 (-1/0/1)
        i32_const(chunk, 0, line);
        chunk.emit_op(i32_cmp_op(&op), line);
      chunk.emit_else(line);

        // both bigint?
        load(chunk, a_slot, line); call1(chunk, test_bigint, line);
        load(chunk, b_slot, line); call1(chunk, test_bigint, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if(line);
          load(chunk, a_slot, line);
          load(chunk, b_slot, line);
          chunk.emit_op(i64_cmp_op(&op), line);
        chunk.emit_else(line);
          // fallback: coerce both to f64
          load(chunk, a_slot, line); call1(chunk, to_f64, line);
          load(chunk, b_slot, line); call1(chunk, to_f64, line);
          chunk.emit_op(f64_cmp_op(&op), line);
        chunk.emit_end(line); // bigint
      chunk.emit_end(line);   // string
    chunk.emit_end(line);     // number
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

    // either is a string → string concatenation
    load(chunk, a_slot, line); call1(chunk, test_str, line);
    load(chunk, b_slot, line); call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_OR, line);        // i32: 1 if either is string
    chunk.emit_if(line);
      // coerce a to string
      load(chunk, a_slot, line);
      call1(chunk, test_str, line);         // i32: 1 if a is already string
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
      load(chunk, a_slot, line); call1(chunk, test_bigint, line);
      load(chunk, b_slot, line); call1(chunk, test_bigint, line);
      chunk.emit_op(Op::I32_AND, line);
      chunk.emit_if(line);
        load(chunk, a_slot, line);
        load(chunk, b_slot, line);
        chunk.emit_op(Op::I64_ADD, line);
      chunk.emit_else(line);
        // number + number (or coerce) → f64.add
        load(chunk, a_slot, line); call1(chunk, to_f64, line);
        load(chunk, b_slot, line); call1(chunk, to_f64, line);
        chunk.emit_op(Op::F64_ADD, line);
        call1(chunk, from_f64, line);
      chunk.emit_end(line); // bigint
    chunk.emit_end(line);   // string
}

// ── emit_dyn_neg ──────────────────────────────────────────────────────

pub fn emit_dyn_neg(chunk: &mut Chunk, line: u32) {
    let v = alloc_locals(chunk, 1);

    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let to_f64      = chunk.add_import("wasm:js-number", "toF64");
    let from_f64    = chunk.add_import("wasm:js-number", "fromF64");

    save(chunk, v, line);

    // bigint → i64 negation
    load(chunk, v, line);
    call1(chunk, test_bigint, line);        // i32: 1 if bigint
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
