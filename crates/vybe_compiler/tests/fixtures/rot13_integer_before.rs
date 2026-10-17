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

pub fn emit_str_rot13(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let rot_slot = alloc_local(chunk);
    let tmp_slot = alloc_local(chunk);

    lset(chunk, s_slot, line);
    emit_parts_new(chunk, line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
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

    // if code >= 65 (possibly a letter)
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(65.0), line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_if(line);
    //   if code <= 90 → uppercase A-Z
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(90.0), line);
    chunk.emit_op(Op::F64_LE, line);
    chunk.emit_if(line);
    emit_rot13_range(chunk, code_slot, 65.0, tmp_slot, line);
    lset(chunk, rot_slot, line);
    chunk.emit_else(line); // code > 90 (still inside code >= 65 branch)
    //   else if code >= 97 → check lowercase a-z
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(97.0), line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_if(line);
    //     if code <= 122 → lowercase a-z
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(122.0), line);
    chunk.emit_op(Op::F64_LE, line);
    chunk.emit_if(line);
    emit_rot13_range(chunk, code_slot, 97.0, tmp_slot, line);
    lset(chunk, rot_slot, line);
    chunk.emit_else(line); // code > 122 → keep
    lget(chunk, code_slot, line);
    lset(chunk, rot_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line); // 91-96 → keep
    lget(chunk, code_slot, line);
    lset(chunk, rot_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line); // end code<=90 if
    chunk.emit_else(line); // code < 65 → keep
    lget(chunk, code_slot, line);
    lset(chunk, rot_slot, line);
    chunk.emit_end(line); // end code>=65 if

    // append chr(rot) to out
    lget(chunk, out_slot, line);
    lget(chunk, rot_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);

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
