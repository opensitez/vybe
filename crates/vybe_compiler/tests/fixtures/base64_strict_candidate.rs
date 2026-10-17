use std::sync::Arc;
use vybe_runtime::{Chunk, Op, Value};
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

fn emit_base64_code_is_allowed(chunk: &mut Chunk, code_slot: u16, line: u32) {
    super::candidate::emit_code_is_allowed(chunk, code_slot, false, line);
}
fn emit_base64_lenient_filter(chunks: &mut [Chunk], current: usize, s_slot: u16, line: u32) {
    lget(&mut chunks[current], s_slot, line);
    vybe_compiler::primitives::base64::emit_filter_lenient(chunks, current, line);
}
pub fn emit_base64_decode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
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
    emit_base64_code_is_allowed(chunk, code_slot, line);
    lset(chunk, ok_slot, line);

    lget(chunk, ok_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(0), line);
    lset(chunk, valid_slot, line);
    chunk.emit_br(loop_state.break_depth(1) as u32, line);
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
