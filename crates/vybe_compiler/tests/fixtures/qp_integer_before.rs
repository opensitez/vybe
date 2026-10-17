use std::sync::Arc;
use vybe_runtime::{Chunk, Op, Value};
fn emit_parts_new(chunk: &mut Chunk, line: u32) {
    let mut imports = Chunk::new("imports");
    vybe_compiler::primitives::collections::emit_array_new_into(&mut imports, chunk, 0, line);
}

fn emit_append_part(chunk: &mut Chunk, parts: u16, line: u32) {
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    lget(chunk, parts, line);
}

fn emit_parts_join(chunk: &mut Chunk, line: u32) {
    push_str(chunk, "", line);
    let join = chunk.add_import("ecma:array", "join");
    chunk.emit_call(join, 2, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}

fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}

fn emit_hexval(chunk: &mut Chunk, code_slot: u16, line: u32) {
    // A decoded digit is numeric; malformed characters return -1.
    for (lo, hi, offset) in [(48.0, 57.0, 48.0), (65.0, 70.0, 55.0), (97.0, 102.0, 87.0)] {
        lget(chunk, code_slot, line);
        chunk.emit_f64_const(lo, line);
        chunk.emit_op(Op::F64_GE, line);
        lget(chunk, code_slot, line);
        chunk.emit_f64_const(hi, line);
        chunk.emit_op(Op::F64_LE, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if_value(line);
        lget(chunk, code_slot, line);
        chunk.emit_f64_const(offset, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_else(line);
    }
    chunk.emit_f64_const(-1.0, line);
    for _ in 0..3 {
        chunk.emit_end(line);
    }
}

pub fn emit_quoted_printable_decode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let c = &mut chunks[current];
    let base = c.alloc_scratch(11);
    let text = base;
    let parts = base + 1;
    let index = base + 2;
    let len = base + 3;
    let part = base + 4;
    let advance = base + 5;
    let look = base + 6;
    let first = base + 7;
    let second = base + 8;
    let hi = base + 9;
    let lo = base + 10;
    lset(c, text, line);
    emit_parts_new(c, line);
    lset(c, parts, line);
    lget(c, text, line);
    let length = c.add_import("wasm:js-string", "length");
    c.emit_call(length, 1, line);
    lset(c, len, line);
    c.emit_f64_const(0.0, line);
    lset(c, index, line);
    let char_at = c.add_import("ecma:string", "charAt");
    let code_at = c.add_import("wasm:js-string", "charCodeAt");
    let from_code = c.add_import("wasm:js-string", "fromCharCode");
    let done = c.emit_block(line);
    let (again, _) = c.emit_loop_s(line);
    lget(c, index, line);
    lget(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    lget(c, text, line);
    lget(c, index, line);
    c.emit_call(char_at, 2, line);
    lset(c, part, line);
    c.emit_f64_const(1.0, line);
    lset(c, advance, line);
    lget(c, text, line);
    lget(c, index, line);
    c.emit_call(code_at, 2, line);
    c.emit_f64_const(61.0, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if(line);
    // Transport padding before a soft break is omitted. All code-unit reads
    // are bounded, including truncated escapes and a final '='.
    lget(c, index, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    lset(c, look, line);
    let padding_done = c.emit_block(line);
    let (padding, _) = c.emit_loop_s(line);
    lget(c, look, line);
    lget(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    lget(c, text, line);
    lget(c, look, line);
    c.emit_call(code_at, 2, line);
    lset(c, first, line);
    for code in [32.0, 9.0] {
        lget(c, first, line);
        c.emit_f64_const(code, line);
        c.emit_op(Op::F64_EQ, line);
    }
    c.emit_op(Op::I32_OR, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    lget(c, look, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    lset(c, look, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(padding);
    c.emit_end(line);
    c.patch_block(padding_done);
    lget(c, look, line);
    lget(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_if(line);
    push_str(c, "", line);
    lset(c, part, line);
    lget(c, len, line);
    lget(c, index, line);
    c.emit_op(Op::F64_SUB, line);
    lset(c, advance, line);
    c.emit_else(line);
    for code in [13.0, 10.0] {
        lget(c, first, line);
        c.emit_f64_const(code, line);
        c.emit_op(Op::F64_EQ, line);
    }
    c.emit_op(Op::I32_OR, line);
    c.emit_if(line);
    push_str(c, "", line);
    lset(c, part, line);
    lget(c, look, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    lset(c, look, line);
    lget(c, first, line);
    c.emit_f64_const(13.0, line);
    c.emit_op(Op::F64_EQ, line);
    lget(c, look, line);
    lget(c, len, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    lget(c, text, line);
    lget(c, look, line);
    c.emit_call(code_at, 2, line);
    c.emit_f64_const(10.0, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if(line);
    lget(c, look, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    lset(c, look, line);
    c.emit_end(line);
    c.emit_end(line);
    lget(c, look, line);
    lget(c, index, line);
    c.emit_op(Op::F64_SUB, line);
    lset(c, advance, line);
    c.emit_else(line);
    lget(c, index, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    lget(c, look, line);
    c.emit_op(Op::F64_EQ, line);
    lget(c, index, line);
    c.emit_f64_const(2.0, line);
    c.emit_op(Op::F64_ADD, line);
    lget(c, len, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    lget(c, text, line);
    lget(c, index, line);
    c.emit_f64_const(2.0, line);
    c.emit_op(Op::F64_ADD, line);
    c.emit_call(code_at, 2, line);
    lset(c, second, line);
    emit_hexval(c, first, line);
    lset(c, hi, line);
    emit_hexval(c, second, line);
    lset(c, lo, line);
    for digit in [hi, lo] {
        lget(c, digit, line);
        c.emit_f64_const(0.0, line);
        c.emit_op(Op::F64_GE, line);
    }
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    lget(c, hi, line);
    c.emit_f64_const(16.0, line);
    c.emit_op(Op::F64_MUL, line);
    lget(c, lo, line);
    c.emit_op(Op::F64_ADD, line);
    c.emit_call(from_code, 1, line);
    lset(c, part, line);
    c.emit_f64_const(3.0, line);
    lset(c, advance, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    lget(c, parts, line);
    lget(c, part, line);
    emit_append_part(c, parts, line);
    c.emit_op(Op::DROP, line);
    lget(c, index, line);
    lget(c, advance, line);
    c.emit_op(Op::F64_ADD, line);
    lset(c, index, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(again);
    c.emit_end(line);
    c.patch_block(done);
    lget(c, parts, line);
    emit_parts_join(c, line);
}
