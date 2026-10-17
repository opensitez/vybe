#![allow(dead_code)]
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

fn emit_uu_div_pow2(chunk: &mut Chunk, slot: u16, div: f64, line: u32) {
    // UU-only: callers supply bounded nonnegative UTF16 units or masked sextets.
    // Divisors are powers of two (4, 16, 64); no byte narrowing is performed.
    lget(chunk, slot, line);
    chunk.emit_i32_const((div as u32).trailing_zeros() as i32, line);
    chunk.emit_op(Op::I32_SHR_U, line);
}

fn emit_uu_mod_pow2(chunk: &mut Chunk, slot: u16, modulus: f64, line: u32) {
    // The same bounded UU integer domain proves this mask equals remainder.
    lget(chunk, slot, line);
    chunk.emit_i32_const(modulus as i32 - 1, line);
    chunk.emit_op(Op::I32_AND, line);
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

pub fn emit_convert_uuencode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let b0_slot = alloc_local(chunk);
    let b1_slot = alloc_local(chunk);
    let b2_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);

    chunk.emit_op_u16(Op::LOCAL_TEE, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);
    // out = fromCharCode(n + 32)  (length char)
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(32.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    let prefix = alloc_local(chunk);
    lset(chunk, prefix, line);
    emit_parts_new(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, out_slot, line);
    lget(chunk, prefix, line);
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    // b0 = charCodeAt(i)
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, b0_slot, line);
    // b1 = (i+1<n) ? charCodeAt(i+1) : 0
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_else(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_end(line);
    lset(chunk, b1_slot, line);
    // b2 = (i+2<n) ? charCodeAt(i+2) : 0
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_else(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_end(line);
    lset(chunk, b2_slot, line);

    // c0 = floor(b0/4)
    emit_uu_div_pow2(chunk, b0_slot, 4.0, line);
    lset(chunk, c_slot, line);
    lget(chunk, out_slot, line);
    emit_uu_enc(chunk, c_slot, line);
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    // c1 = (b0 mod 4)*16 + floor(b1/16)
    emit_uu_mod_pow2(chunk, b0_slot, 4.0, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_uu_div_pow2(chunk, b1_slot, 16.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, c_slot, line);
    lget(chunk, out_slot, line);
    emit_uu_enc(chunk, c_slot, line);
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    // c2 = (b1 mod 16)*4 + floor(b2/64)
    emit_uu_mod_pow2(chunk, b1_slot, 16.0, line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_uu_div_pow2(chunk, b2_slot, 64.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, c_slot, line);
    lget(chunk, out_slot, line);
    emit_uu_enc(chunk, c_slot, line);
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    // c3 = b2 mod 64
    emit_uu_mod_pow2(chunk, b2_slot, 64.0, line);
    lset(chunk, c_slot, line);
    lget(chunk, out_slot, line);
    emit_uu_enc(chunk, c_slot, line);
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(3.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    push_str(chunk, "\n", line);
    emit_append_part(chunk, out_slot, line);
    emit_parts_join(chunk, line);
}

pub fn emit_convert_uudecode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let count_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    let done_slot = alloc_local(chunk);
    let c0 = alloc_local(chunk);
    let c1 = alloc_local(chunk);
    let c2 = alloc_local(chunk);
    let c3 = alloc_local(chunk);
    let byte_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);

    lset(chunk, s_slot, line);
    emit_parts_new(chunk, line);
    lset(chunk, out_slot, line);
    // Bound the prefix read and the full first-line payload. A malformed
    // buffer returns false rather than trapping in a string builtin.
    lget(chunk, s_slot, line);
    let length = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(length, 1, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, len_slot, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_end(line);
    // The prefix and payload both carry six-bit values.
    lget(chunk, s_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, count_slot, line);
    emit_uu_dec(chunk, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, count_slot, line);
    chunk.emit_f64_const(3.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_CEIL, line);
    chunk.emit_f64_const(4.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_end(line);
    push_const(chunk, Value::F64(1.0), line);
    lset(chunk, pos_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, done_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, done_slot, line);
    lget(chunk, count_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    // decode c0..c3 from four chars at pos
    for (k, slot) in [(0.0, c0), (1.0, c1), (2.0, c2), (3.0, c3)] {
        lget(chunk, s_slot, line);
        lget(chunk, pos_slot, line);
        push_const(chunk, Value::F64(k), line);
        chunk.emit_op(Op::F64_ADD, line);
        {
            let idx = chunk.add_import("wasm:js-string", "charCodeAt");
            chunk.emit_call(idx, 2, line);
        }
        let tmp = alloc_local(chunk);
        lset(chunk, tmp, line);
        emit_uu_dec(chunk, tmp, line);
        lset(chunk, slot, line);
    }
    // b0 = c0*4 + floor(c1/16)
    lget(chunk, c0, line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_uu_div_pow2(chunk, c1, 16.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, byte_slot, line);
    lget(chunk, out_slot, line);
    lget(chunk, byte_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    lget(chunk, done_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, done_slot, line);

    // if done < count: emit b1
    lget(chunk, done_slot, line);
    lget(chunk, count_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if(line);
    // b1 = (c1 mod 16)*16 + floor(c2/4)
    emit_uu_mod_pow2(chunk, c1, 16.0, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_uu_div_pow2(chunk, c2, 4.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, byte_slot, line);
    lget(chunk, out_slot, line);
    lget(chunk, byte_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    lget(chunk, done_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, done_slot, line);
    chunk.emit_end(line);

    // if done < count: emit b2
    lget(chunk, done_slot, line);
    lget(chunk, count_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if(line);
    // b2 = (c2 mod 4)*64 + c3
    emit_uu_mod_pow2(chunk, c2, 4.0, line);
    push_const(chunk, Value::F64(64.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    lget(chunk, c3, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, byte_slot, line);
    lget(chunk, out_slot, line);
    lget(chunk, byte_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    lget(chunk, done_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, done_slot, line);
    chunk.emit_end(line);

    lget(chunk, pos_slot, line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, pos_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    emit_parts_join(chunk, line);
}
