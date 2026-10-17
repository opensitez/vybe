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

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
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

pub fn emit_quoted_printable_encode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let hex_slot = alloc_local(chunk);

    lset(chunk, s_slot, line);
    push_str(chunk, "0123456789ABCDEF", line);
    lset(chunk, hex_slot, line);
    emit_parts_new(chunk, line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_LT, line);
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

    // charCodeAt proves0..65535. Outside printable ASCII, only TAB is
    // exempt; '=' is the sole printable unit that requires escaping.
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_i32_const(94, line);
    chunk.emit_op(Op::I32_GT_U, line);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(61, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_op(Op::I32_OR, line);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(9, line);
    chunk.emit_op(Op::I32_NE, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    // Keep the original high index unmasked: units above255 retain their
    // established empty high-digit charAt result instead of being narrowed.
    lget(chunk, out_slot, line);
    push_str(chunk, "=", line);
    emit_append_part(chunk, out_slot, line);
    lget(chunk, hex_slot, line);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    emit_append_part(chunk, out_slot, line);
    lget(chunk, hex_slot, line);
    // For the full unsigned UTF-16 domain, remainder16 is the low nibble.
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(15, line);
    chunk.emit_op(Op::I32_AND, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    emit_parts_join(chunk, line);
}
