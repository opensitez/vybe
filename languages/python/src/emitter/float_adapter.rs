//! Python float repr.
//!
//! Formats a value the walker has *statically* determined to be a float so a
//! whole number prints with a trailing `.0` (`4.0`, not `4`) and `inf`/`nan`
//! use Python's spellings. This is only invoked on expressions known to be
//! floats (float literals, `/` true division, `float()`, float-returning math)
//! — never a blanket cast of every number.

use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::ops;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const FLOAT_FIELDS_KEY: &str = "__py_float_fields";

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn stash_exact(chunks: &mut [Chunk], current: usize, argc: u8, want: u16, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(want.max(1));
    for offset in (want..argc as u16).rev() {
        let _ = offset;
        chunks[current].emit_op(Op::DROP, line);
    }
    for offset in (0..want).rev() {
        if offset < argc as u16 {
            lset(&mut chunks[current], base + offset, line);
        } else {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            lset(&mut chunks[current], base + offset, line);
        }
    }
    base
}

fn emit_slot_is_null_or_undefined(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_undefined = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(is_undefined, 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

/// Python float display. Stack: `[num]` → `[string]`.
pub fn emit_float_repr(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let x = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);

    // isNaN(x) → "nan"
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let is_nan = chunk.add_import("ecma:number", "isNaN");
    chunk.emit_call(is_nan, 1, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("nan", line);
    chunk.emit_else(line);

    // !isFinite(x) → "inf" / "-inf"
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let is_finite = chunk.add_import("ecma:number", "isFinite");
    chunk.emit_call(is_finite, 1, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    ops::emit_dyn_lt(chunk, line); // 0 < x ?
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("inf", line);
    chunk.emit_else(line);
    chunk.emit_string_const("-inf", line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    // finite: whole number → String(x) + ".0", else String(x)
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let is_int = chunk.add_import("ecma:number", "isInteger");
    chunk.emit_call(is_int, 1, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    // `String(-0)` is "0" per ECMA-262 §6.1.6.1.20, but CPython reprs negative
    // zero as "-0.0". The sign survives in the f64, so read it back with
    // `copysign(1, x) < 0` — true for -0.0 as well as ordinary negatives (for
    // which `String` already emits the sign, so only zero needs the prefix).
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    chunk.emit_call(to_f64, 1, line);
    core_wasm::f64_const(chunk, line, 0.0);
    chunk.emit_op(Op::F64_EQ, line);
    // `f64.copysign(a, b)` = magnitude of `a`, sign of `b` — so 1.0 is pushed
    // FIRST and x second, giving ±1 carrying x's sign bit.
    core_wasm::f64_const(chunk, line, 1.0);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_op(Op::F64_COPYSIGN, line);
    core_wasm::f64_const(chunk, line, 0.0);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("-", line);
    chunk.emit_else(line);
    chunk.emit_string_const("", line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let to_str = chunk.add_import("ecma:string", "String");
    chunk.emit_call(to_str, 1, line);
    let concat = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(concat, 2, line);
    chunk.emit_string_const(".0", line);
    chunk.emit_call(concat, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let to_str2 = chunk.add_import("ecma:string", "String");
    chunk.emit_call(to_str2, 1, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Stamp an object with the fields whose runtime values should use Python
/// float display. Stack: `[obj, fields]` -> `[obj]`.
pub fn emit_stamp_float_fields(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let obj = base;
    let fields = base + 1;
    lget(&mut chunks[current], obj, line);
    lget(&mut chunks[current], fields, line);
    let slot = class_slots::resolve(&ClassSlot::internal(FLOAT_FIELDS_KEY), &PlainNames);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Stack,
        &slot,
        ValueSource::Stack,
        line,
    );
    lget(&mut chunks[current], obj, line);
}

/// Format `obj[field]`, consulting the object's float-field stamp. Stack:
/// `[obj, field]` -> `[string]`.
pub fn emit_float_field_str(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let obj = base;
    let field = base + 1;
    let value = chunks[current].alloc_scratch(1);
    let fields = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], obj, line);
    lget(&mut chunks[current], field, line);
    let get = chunks[current].add_import("ecma:object", "get");
    chunks[current].emit_call(get, 2, line);
    lset(&mut chunks[current], value, line);

    let slot = class_slots::resolve(&ClassSlot::internal(FLOAT_FIELDS_KEY), &PlainNames);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(obj),
        &slot,
        Dest::Local(fields),
        line,
    );
    emit_slot_is_null_or_undefined(&mut chunks[current], fields, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value, line);
    let to_str = chunks[current].add_import("ecma:string", "String");
    chunks[current].emit_call(to_str, 1, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], fields, line);
    lget(&mut chunks[current], field, line);
    let includes = chunks[current].add_import("ecma:array", "includes");
    chunks[current].emit_call(includes, 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value, line);
    emit_float_repr(chunks, current, 1, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value, line);
    let to_str = chunks[current].add_import("ecma:string", "String");
    chunks[current].emit_call(to_str, 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}
