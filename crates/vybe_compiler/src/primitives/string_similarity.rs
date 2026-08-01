//! `string_similarity` — string distance and phonetic-key adapter primitives.
//!
//! ECMA-262 has no equivalent for any of these, so every language that offers
//! them has had to implement them itself. They are pure, language-neutral
//! algorithms over text — PHP `levenshtein`/`similar_text`/`soundex`/
//! `metaphone`, Python `difflib` ratios, Java/Ruby user code — so they live
//! here once rather than in each `string_adapter.rs`.
//!
//! **No coercion here.** Every function takes STRINGS on the stack. How a
//! language turns a non-string into a string is the language's own rule and
//! stays at its call site — PHP's `trim(true)` is `"1"` where ECMA's
//! `String(true)` is `"true"`, and Python raises instead of coercing.

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
        Value::Null => chunk.emit_op(Op::NULL, line),
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

fn emit_soundex_digit(
    chunks: &mut [Chunk],
    current: usize,
    code_slot: u16,
    digit_slot: u16,
    line: u32,
) {
    // Range table: A→default 0, then per character.
    // BFPV → "1"; CGJKQSXZ → "2"; DT → "3"; L → "4"; MN → "5"; R → "6"
    // Everything else (vowels, H, W, Y, non-letters) → "0"
    // Implementation: a long if-else chain by char code.
    let table: &[(&[u32], &str)] = &[
        (&[66, 70, 80, 86], "1"),
        (&[67, 71, 74, 75, 81, 83, 88, 90], "2"),
        (&[68, 84], "3"),
        (&[76], "4"),
        (&[77, 78], "5"),
        (&[82], "6"),
    ];
    let chunk = &mut chunks[current];
    push_str(chunk, "0", line);
    lset(chunk, digit_slot, line);
    for (codes, digit) in table {
        for &cc in *codes {
            lget(chunk, code_slot, line);
            push_const(chunk, Value::F64(cc as f64), line);
            crate::primitives::ops::emit_dyn_eq(chunk, line);
            crate::primitives::ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_if(line);
            push_str(chunk, digit, line);
            lset(chunk, digit_slot, line);
            chunk.emit_end(line);
        }
    }
}

pub fn emit_soundex(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let last_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let digit_slot = alloc_local(chunk);

    {
        let idx = chunk.add_import("ecma:string", "toUpperCase");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);

    // if s.length == 0: return ""; else: compute soundex
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "", line);
    chunk.emit_else(line);

    // out = first letter
    lget(chunk, s_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, out_slot, line);

    // last = digit_for(first letter)
    push_str(chunk, "0", line);
    lset(chunk, last_slot, line);
    lget(chunk, s_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);
    emit_soundex_digit(chunks, current, code_slot, digit_slot, line);
    let chunk = &mut chunks[current];
    lget(chunk, digit_slot, line);
    lset(chunk, last_slot, line);

    // i = 1; n = s.length
    push_const(chunk, Value::F64(1.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);

    // while i < n && out.length < 4
    let _ = chunk;
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    // compound condition: i < n AND out.length < 4
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, out_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(4.0), line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // c = s.charAt(i); code = s.charCodeAt(i); digit = lookup
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, c_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);
    emit_soundex_digit(chunks, current, code_slot, digit_slot, line);
    let chunk = &mut chunks[current];

    // if digit != "0" and digit != last: out += digit; last = digit
    // else: last = "0" (vowels reset last)
    lget(chunk, digit_slot, line);
    push_str(chunk, "0", line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // digit == "0": reset last
    push_str(chunk, "0", line);
    lset(chunk, last_slot, line);
    chunk.emit_else(line);
    // digit != "0": check if same as last; if not same: append
    lget(chunk, digit_slot, line);
    lget(chunk, last_slot, line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    crate::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);
    // different non-zero: append
    lget(chunk, out_slot, line);
    lget(chunk, digit_slot, line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, digit_slot, line);
    lset(chunk, last_slot, line);
    chunk.emit_end(line); // end not_same if
    chunk.emit_end(line); // end digit=0 if

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    // pad out with "0" until length 4
    let _ = chunk;
    let pad_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(4.0), line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    push_str(chunk, "0", line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, pad_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
    chunk.emit_end(line); // end empty check if/else
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
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, prev_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    let init_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
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
    crate::primitives::loops::emit_loop_end(chunks, current, init_state, line);

    // curr = new array of n+1 zeros
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, curr_slot, line);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
    let init2_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
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
    crate::primitives::loops::emit_loop_end(chunks, current, init2_state, line);

    // Outer loop: for i in 1..=m
    push_const(&mut chunks[current], Value::F64(1.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let outer_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, m_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
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
    let inner_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
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
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    crate::primitives::loops::emit_loop_end(chunks, current, inner_state, line);
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
    crate::primitives::loops::emit_loop_end(chunks, current, outer_state, line);
    let chunk = &mut chunks[current];

    // Result: prev[n]
    lget(chunk, prev_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

pub fn emit_similar_text(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let b_slot = alloc_local(chunk);
    let a_slot = alloc_local(chunk);
    let used_slot = alloc_local(chunk);
    let total_slot = alloc_local(chunk);
    let m_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let j_slot = alloc_local(chunk);

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

    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, used_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, total_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    let init_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, used_slot, line);
    push_const(chunk, Value::Bool(false), line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, init_state, line);

    push_const(&mut chunks[current], Value::F64(0.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let outer_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, m_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);

    push_const(&mut chunks[current], Value::F64(0.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
    let inner_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // if !used[j] && a[i] == b[j]: mark used, count, break inner
    lget(chunk, used_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    crate::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line); // if !used[j]

    lget(chunk, a_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lget(chunk, b_slot, line);
    lget(chunk, j_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line); // if a[i]==b[j]

    lget(chunk, used_slot, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::Bool(true), line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];

    lget(chunk, total_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, total_slot, line);
    // break out of inner loop (depth 1 = inner block)
    chunk.emit_br(inner_state.break_depth(0) as u32, line);

    chunk.emit_end(line); // end if a[i]==b[j]
    chunk.emit_end(line); // end if !used[j]

    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, inner_state, line);

    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, outer_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, total_slot, line);
}

pub fn emit_metaphone(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // Optional max-phonemes arg (TOS). PHP: 0 means "no limit".
    let limit_slot = alloc_local(chunk);
    if argc >= 2 {
        crate::primitives::convert::emit_to_int(chunk, line);
        lset(chunk, limit_slot, line);
    }
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);

    {
        let idx = chunk.add_import("ecma:string", "toUpperCase");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
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
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, c_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    // First letter always kept; subsequent: skip vowels/H/W/Y, keep consonants
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // first: append c
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_else(line);

    // check is_vowel_or_h: code in {65=A, 69=E, 73=I, 79=O, 85=U, 72=H, 87=W, 89=Y}
    // Compute is_vowel as a local flag
    let is_vowel_slot = alloc_local(chunk);
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, is_vowel_slot, line);
    for &cc in &[65u32, 69, 73, 79, 85, 72, 87, 89] {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::F64(cc as f64), line);
        crate::primitives::ops::emit_dyn_eq(chunk, line);
        crate::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::Bool(true), line);
        lset(chunk, is_vowel_slot, line);
        chunk.emit_end(line);
    }
    // if !is_vowel: append if code in A-Z range
    lget(chunk, is_vowel_slot, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    crate::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);
    // not a vowel: append if 65 <= code <= 90
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(65.0), line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line); // >= 65
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(90.0), line);
    crate::primitives::ops::emit_dyn_gt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line); // <= 90
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line); // end range check
    chunk.emit_end(line); // end not_vowel if
    chunk.emit_end(line); // end first/subsequent if

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    // Apply the max-phonemes limit (only when > 0; PHP treats 0 as no limit).
    if argc >= 2 {
        lget(chunk, limit_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        crate::primitives::ops::emit_dyn_gt(chunk, line);
        crate::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        lget(chunk, out_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        lget(chunk, limit_slot, line);
        crate::primitives::strings::emit_substring(chunk, line);
        chunk.emit_else(line);
        lget(chunk, out_slot, line);
        chunk.emit_end(line);
    } else {
        lget(chunk, out_slot, line);
    }
}
