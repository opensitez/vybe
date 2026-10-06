// Exact baseline helpers and selected base64 decode body, public guard removed only in fixture.
use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// All output fragments below are proven strings (literals, charAt, or
// fromCharCode). Collect them privately and join once, avoiding quadratic
// copying and dynamic addition/coercion dispatch on every append.
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

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
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

#[allow(dead_code)]
fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, argc, line);
}

fn emit_base64_lenient_filter(chunks: &mut [Chunk], current: usize, s_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let ok_slot = alloc_local(chunk);

    emit_parts_new(chunk, line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::I32(0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::I32_LT_S, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    push_const(chunk, Value::I32(0), line);
    lset(chunk, ok_slot, line);
    for (lo, hi) in [(65, 90), (97, 122), (48, 57)] {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::I32(lo), line);
        chunk.emit_op(Op::I32_GE_S, line);
        lget(chunk, code_slot, line);
        push_const(chunk, Value::I32(hi), line);
        chunk.emit_op(Op::I32_LE_S, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if(line);
        push_const(chunk, Value::I32(1), line);
        lset(chunk, ok_slot, line);
        chunk.emit_end(line);
    }
    for code in [43, 47, 61] {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::I32(code), line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if(line);
        push_const(chunk, Value::I32(1), line);
        lset(chunk, ok_slot, line);
        chunk.emit_end(line);
    }

    lget(chunk, ok_slot, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lget(chunk, code_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    emit_parts_join(chunk, line);
}

fn emit_floordiv(chunk: &mut Chunk, slot: u16, div: f64, line: u32) {
    lget(chunk, slot, line);
    push_const(chunk, Value::F64(div), line);
    chunk.emit_op(Op::F64_DIV, line);
    vybe_compiler::primitives::math::emit_floor(chunk, line);
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

fn emit_modn(chunk: &mut Chunk, slot: u16, m: f64, line: u32) {
    lget(chunk, slot, line);
    emit_floordiv(chunk, slot, m, line);
    push_const(chunk, Value::F64(m), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
}

fn emit_rot13_range(chunk: &mut Chunk, code_slot: u16, base: f64, tmp_slot: u16, line: u32) {
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(base), line);
    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::F64(13.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, tmp_slot, line);
    // if tmp >= 26: tmp -= 26
    lget(chunk, tmp_slot, line);
    push_const(chunk, Value::F64(26.0), line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_if(line);
    lget(chunk, tmp_slot, line);
    push_const(chunk, Value::F64(26.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, tmp_slot, line);
    chunk.emit_end(line);
    lget(chunk, tmp_slot, line);
    push_const(chunk, Value::F64(base), line);
    chunk.emit_op(Op::F64_ADD, line);
}

fn emit_uu_dec(chunk: &mut Chunk, code_slot: u16, line: u32) {
    lget(chunk, code_slot, line);
    chunk.emit_f64_const(32.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_i32_const(63, line);
    chunk.emit_op(Op::I32_AND, line);
}

fn emit_uu_enc(chunk: &mut Chunk, c_slot: u16, line: u32) {
    lget(chunk, c_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(96.0), line);
    chunk.emit_else(line);
    lget(chunk, c_slot, line);
    push_const(chunk, Value::F64(32.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_end(line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
}

pub fn emit_body(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Absence of a strict argument proves the lenient default. Do not emit
    // strict validation or a dynamic truthiness ladder for a known default.
    if argc < 2 {
        let input = alloc_local(&mut chunks[current]);
        lset(&mut chunks[current], input, line);
        emit_base64_lenient_filter(chunks, current, input, line);
        vybe_compiler::primitives::base64::emit_decode_binary_string(chunks, current, line);
        return;
    }

    let chunk = &mut chunks[current];
    let strict_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let valid_slot = alloc_local(chunk);

    if argc >= 2 {
        lset(chunk, strict_slot, line);
    } else {
        push_const(chunk, Value::Bool(false), line);
        lset(chunk, strict_slot, line);
    }
    lset(chunk, s_slot, line);

    lget(chunk, strict_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    push_const(chunk, Value::I32(1), line);
    lset(chunk, valid_slot, line);
    push_const(chunk, Value::I32(0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::I32_LT_S, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    let ok_slot = alloc_local(chunk);
    push_const(chunk, Value::I32(0), line);
    lset(chunk, ok_slot, line);

    for (lo, hi) in [(65, 90), (97, 122), (48, 57)] {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::I32(lo), line);
        chunk.emit_op(Op::I32_GE_S, line);
        lget(chunk, code_slot, line);
        push_const(chunk, Value::I32(hi), line);
        chunk.emit_op(Op::I32_LE_S, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if(line);
        push_const(chunk, Value::I32(1), line);
        lset(chunk, ok_slot, line);
        chunk.emit_end(line);
    }
    for code in [43, 47, 61] {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::I32(code), line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if(line);
        push_const(chunk, Value::I32(1), line);
        lset(chunk, ok_slot, line);
        chunk.emit_end(line);
    }

    lget(chunk, ok_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(0), line);
    lset(chunk, valid_slot, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, valid_slot, line);
    chunk.emit_if_value(line);
    lget(chunk, s_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::base64::emit_decode_binary_string(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    let _ = chunk;
    emit_base64_lenient_filter(chunks, current, s_slot, line);
    let chunk = &mut chunks[current];
    let _ = chunk;
    vybe_compiler::primitives::base64::emit_decode_binary_string(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
}
