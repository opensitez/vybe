use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

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

pub fn emit_levenshtein(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = alloc_local(chunk);
    let a_slot = alloc_local(chunk);
    let m_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let prev_slot = alloc_local(chunk);
    let curr_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let j_slot = alloc_local(chunk);
    let cost_slot = alloc_local(chunk);
    let tmp_slot = alloc_local(chunk);
    let v_slot = alloc_local(chunk);

    lset(chunk, b_slot, line);
    lset(chunk, a_slot, line);

    lget(chunk, a_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, m_slot, line);
    lget(chunk, b_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);

    // prev[j] = j  (distance from "" to b[..j])
    chunk.emit_array_new_fixed(0, 0, line);
    lset(chunk, prev_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    let init_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, prev_slot, line);
    lget(chunk, j_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, init_state, line);

    // curr = new array of n+1 zeros
    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, curr_slot, line);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
    let init2_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, curr_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, init2_state, line);

    // Outer loop: for i in 1..=m
    push_const(&mut chunks[current], Value::F64(1.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let outer_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, m_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    // curr[0] = i
    lget(chunk, curr_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, i_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    push_const(&mut chunks[current], Value::F64(1.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
    let inner_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    // cost = (a[i-1] == b[j-1]) ? 0 : 1
    lget(chunk, a_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lget(chunk, b_slot, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    emit_str_equals(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_end(line);
    lset(chunk, cost_slot, line);

    // del = curr[j-1] + 1
    lget(chunk, curr_slot, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, tmp_slot, line);

    // ins = prev[j] + 1
    lget(chunk, prev_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, v_slot, line);

    // tmp = min(tmp, v)
    lget(chunk, v_slot, line);
    lget(chunk, tmp_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if(line);
    lget(chunk, v_slot, line);
    lset(chunk, tmp_slot, line);
    chunk.emit_end(line);

    // sub = prev[j-1] + cost
    lget(chunk, prev_slot, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lget(chunk, cost_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, v_slot, line);

    // tmp = min(tmp, v)
    lget(chunk, v_slot, line);
    lget(chunk, tmp_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if(line);
    lget(chunk, v_slot, line);
    lset(chunk, tmp_slot, line);
    chunk.emit_end(line);

    // curr[j] = tmp
    lget(chunk, curr_slot, line);
    lget(chunk, j_slot, line);
    lget(chunk, tmp_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];

    // j++
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, inner_state, line);
    let chunk = &mut chunks[current];

    // swap prev <-> curr
    lget(chunk, prev_slot, line);
    lset(chunk, tmp_slot, line);
    lget(chunk, curr_slot, line);
    lset(chunk, prev_slot, line);
    lget(chunk, tmp_slot, line);
    lset(chunk, curr_slot, line);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer_state, line);
    let chunk = &mut chunks[current];

    // Result: prev[n]
    lget(chunk, prev_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

// Captured private helper: the same string-equality import and arity.
fn emit_str_equals(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "equals");
    chunk.emit_call(idx, 2, line);
}
