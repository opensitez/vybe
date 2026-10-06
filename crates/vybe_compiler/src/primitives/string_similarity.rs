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
//!
//! The algorithm's private counters/distances are numbers, character results
//! and phonetic digits are strings, and flags are booleans. Their producers
//! establish the type: compare them directly instead of emitting dynamic
//! equality, arithmetic, and truthiness dispatch inside the nested loops.

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

fn emit_soundex_digit(
    chunks: &mut [Chunk],
    current: usize,
    code_slot: u16,
    digit_slot: u16,
    line: u32,
) {
    // A..Z -> the same Soundex groups, as a constant indexed string. The
    // unsigned bounds test also rejects codes below A after subtraction.
    let chunk = &mut chunks[current];
    let index = alloc_local(chunk);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(65, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, index, line);
    chunk.emit_i32_const(25, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_if_value(line);
    push_str(chunk, "01230120022455012623010202", line);
    lget(chunk, index, line);
    lget(chunk, index, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    crate::primitives::strings::emit_str_substring(chunk, line);
    chunk.emit_else(line);
    push_str(chunk, "0", line);
    chunk.emit_end(line);
    lset(chunk, digit_slot, line);
}

pub fn emit_soundex(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let last_slot = alloc_local(chunk);
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
    chunk.emit_op(Op::F64_EQ, line);
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
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_i32(line);
    lget(chunk, out_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    // code = s.charCodeAt(i); digit = lookup
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
    crate::primitives::strings::emit_str_equals(chunk, line);
    chunk.emit_if(line);
    // digit == "0": reset last
    push_str(chunk, "0", line);
    lset(chunk, last_slot, line);
    chunk.emit_else(line);
    // digit != "0": check if same as last; if not same: append
    lget(chunk, digit_slot, line);
    lget(chunk, last_slot, line);
    crate::primitives::strings::emit_str_equals(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    // different non-zero: append
    lget(chunk, out_slot, line);
    lget(chunk, digit_slot, line);
    crate::primitives::strings::emit_concat(chunk, 2, line);
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
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    push_str(chunk, "0", line);
    crate::primitives::strings::emit_concat(chunk, 2, line);
    lset(chunk, out_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, pad_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
    chunk.emit_end(line); // end empty check if/else
}

pub fn emit_levenshtein(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    crate::primitives::bundle::emit_private_helper_call(
        &mut chunks[current],
        "__stdlib_levenshtein",
        2,
        line,
    );
}

pub(crate) fn build_levenshtein() -> Chunk {
    let mut chunk = crate::primitives::functions::create_function_chunk("__stdlib_levenshtein", 2);
    chunk.local_count = 2;
    crate::primitives::bundle::mark_private_helper_body(&mut chunk);
    let mut chunks = vec![chunk];
    emit_levenshtein_body(&mut chunks, 0, 0, 1, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    chunks.remove(0)
}

pub(crate) fn emit_levenshtein_body(
    chunks: &mut [Chunk],
    current: usize,
    a_slot: u16,
    b_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let m_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let prev_slot = alloc_local(chunk);
    let curr_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let j_slot = alloc_local(chunk);
    let cost_slot = alloc_local(chunk);
    let tmp_slot = alloc_local(chunk);
    let v_slot = alloc_local(chunk);
    let a_char_slot = alloc_local(chunk);
    let b_chars_slot = alloc_local(chunk);

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

    // Strings are immutable. Preserve charAt's UTF-16/code-unit contract,
    // but calculate each b character once rather than in every matrix row.
    chunk.emit_array_new_fixed(0, 0, line);
    lset(chunk, b_chars_slot, line);
    lget(chunk, m_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if(line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    let cache_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, b_chars_slot, line);
    lget(chunk, b_slot, line);
    lget(chunk, j_slot, line);
    let idx = chunk.add_import("ecma:string", "charAt");
    chunk.emit_call(idx, 2, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, cache_state, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);

    // prev[j] = j  (distance from "" to b[..j])
    chunk.emit_array_new_fixed(0, 0, line);
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
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
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
    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, curr_slot, line);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
    let init2_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
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
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    // curr[0] = i
    lget(chunk, curr_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, i_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    let chunk = &mut chunks[current];
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if(line);
    lget(chunk, a_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, a_char_slot, line);
    chunk.emit_end(line);

    push_const(&mut chunks[current], Value::F64(1.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
    let inner_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    // cost = (a[i-1] == b[j-1]) ? 0 : 1
    lget(chunk, a_char_slot, line);
    lget(chunk, b_chars_slot, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    crate::primitives::strings::emit_str_equals(chunk, line);
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

pub fn emit_similar_text(chunks:&mut [Chunk],current:usize,argc:u8,line:u32) {
 for _ in 2..argc {chunks[current].emit_op(Op::DROP,line);}
 crate::primitives::bundle::emit_private_helper_call(&mut chunks[current],"__stdlib_similar_text",2,line);
}
pub(crate) fn build_similar_text()->Chunk {
 let mut chunk=crate::primitives::functions::create_function_chunk("__stdlib_similar_text",2);
 chunk.local_count=2;crate::primitives::bundle::mark_private_helper_body(&mut chunk);
 let mut chunks=vec![chunk];emit_similar_text_body(&mut chunks,0,0,1,0);
 chunks[0].emit_op(Op::RETURN,0);chunks.remove(0)
}
pub(crate) fn emit_similar_text_body(chunks:&mut [Chunk],current:usize,a:u16,b:u16,line:u32) {
    similar_text_algorithm::emit_body(chunks,current,a,b,line);
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
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
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
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);
    // first: append c
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    crate::primitives::strings::emit_concat(chunk, 2, line);
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
        chunk.emit_op(Op::F64_EQ, line);
        chunk.emit_if(line);
        push_const(chunk, Value::Bool(true), line);
        lset(chunk, is_vowel_slot, line);
        chunk.emit_end(line);
    }
    // if !is_vowel: append if code in A-Z range
    lget(chunk, is_vowel_slot, line);
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    // not a vowel: append if 65 <= code <= 90
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(65.0), line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_op(Op::I32_EQZ, line); // >= 65
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(90.0), line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_op(Op::I32_EQZ, line); // <= 90
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    crate::primitives::strings::emit_concat(chunk, 2, line);
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
        chunk.emit_op(Op::F64_GT, line);
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

mod similar_text_algorithm {
//! Separate frequency candidate with compliant private-buffer fallback.
use vybe_runtime::{Chunk, Op};
mod baseline {
//! Isolated standard-GC private-buffer foundation; frozen baseline untouched.
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

// Exact frozen private helper expansion, not a test-only dispatcher.
fn emit_str_equals(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "equals");
    chunk.emit_call(idx, 2, line);
}

/// Inputs/lengths already strict-validated by caller; exact remaining loop.
pub fn emit_prevalidated(chunks:&mut [Chunk],current:usize,line:u32,
 b_slot:u16,a_slot:u16,used_slot:u16,total_slot:u16,m_slot:u16,n_slot:u16,i_slot:u16,j_slot:u16) {
 let chunk=&mut chunks[current];
    // Separate compliant-buffer foundation, not frozen preimage or production.
    // Original only observes positions j<n; trailing n+1 position is unused.
    lget(chunk,n_slot,line);
    chunk.emit_i32_const(0,line);
    chunk.emit_op(Op::I32_LT_S,line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(0,line); // preserve signed-negative original loop behavior
    chunk.emit_else(line);lget(chunk,n_slot,line);chunk.emit_end(line);
    chunk.emit_op_u16(Op::ARRAY_NEW_DEFAULT,0,line);
    lset(chunk,used_slot,line);
    push_const(chunk,Value::F64(0.),line);lset(chunk,total_slot,line);
    push_const(chunk,Value::F64(0.),line);lset(chunk,j_slot,line);
    let _=chunk;
    // Null default lanes are private and never escape before false filling.
    let init_state=crate::primitives::loops::emit_loop_start(chunks,current,line);
    let chunk=&mut chunks[current];lget(chunk,j_slot,line);lget(chunk,n_slot,line);
    chunk.emit_op(Op::F64_LT,line);let _=chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks,current,line);
    let chunk=&mut chunks[current];lget(chunk,used_slot,line);lget(chunk,j_slot,line);
    push_const(chunk,Value::Bool(false),line);chunk.emit_op(Op::ARRAY_SET,line);
    lget(chunk,j_slot,line);push_const(chunk,Value::F64(1.),line);
    chunk.emit_op(Op::F64_ADD,line);lset(chunk,j_slot,line);let _=chunk;
    crate::primitives::loops::emit_loop_end(chunks,current,init_state,line);

    push_const(&mut chunks[current], Value::F64(0.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let outer_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, m_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);

    push_const(&mut chunks[current], Value::F64(0.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
    let inner_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    // if !used[j] && a[i] == b[j]: mark used, count, break inner
    lget(chunk, used_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
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
    emit_str_equals(chunk, line);
    chunk.emit_if(line); // if a[i]==b[j]

    lget(chunk, used_slot, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_op(Op::ARRAY_SET,line);
    let _=chunk;
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

}
fn get(c:&mut Chunk,s:u16,l:u32){c.emit_op_u16(Op::LOCAL_GET,s,l);}
fn set(c:&mut Chunk,s:u16,l:u32){c.emit_op_u16(Op::LOCAL_SET,s,l);}
fn call(c:&mut Chunk,m:&str,n:&str,a:u8,l:u32){let i=c.add_import(m,n);c.emit_call(i,a,l);}
fn flag_if(c:&mut Chunk,s:u16,l:u32){get(c,s,l);call(c,"wasm:js-number","toI32",1,l);c.emit_if(l);}
fn increment(c:&mut Chunk,s:u16,l:u32){get(c,s,l);c.emit_f64_const(1.,l);c.emit_op(Op::F64_ADD,l);set(c,s,l);}
fn loop_begin(c:&mut Chunk,index:u16,len:u16,l:u32)->(usize,usize){
    c.emit_f64_const(0.,l);set(c,index,l);
    let block=c.emit_block(l);let (lp,_)=c.emit_loop_s(l);
    get(c,index,l);get(c,len,l);c.emit_op(Op::F64_LT,l);c.emit_op(Op::I32_EQZ,l);c.emit_br_if(1,l);
    (block,lp)
}
fn loop_end(c:&mut Chunk,index:u16,patch:(usize,usize),l:u32){
    increment(c,index,l);c.emit_br(0,l);c.emit_end(l);c.patch_loop(patch.1);c.emit_end(l);c.patch_block(patch.0);
}
fn code(c:&mut Chunk,text:u16,index:u16,l:u32){get(c,text,l);get(c,index,l);call(c,"wasm:js-string","charCodeAt",2,l);}
pub(super) fn emit_body(chunks:&mut [Chunk],current:usize,a:u16,b:u16,line:u32) {
    let c=&mut chunks[current];let base=c.local_count;
    let used=c.alloc_scratch(1);let total=c.alloc_scratch(1);
    let m=c.alloc_scratch(1);let n=c.alloc_scratch(1);
    let i=c.alloc_scratch(1);let j=c.alloc_scratch(1);
    get(c,a,line);call(c,"wasm:js-string","length",1,line);set(c,m,line);
    get(c,b,line);call(c,"wasm:js-string","length",1,line);set(c,n,line);
    // Experimental threshold, report its small/cold costs independently.
    get(c,m,line);get(c,n,line);c.emit_op(Op::F64_MUL,line);
    c.emit_f64_const(4096.,line);c.emit_op(Op::F64_GE,line);set(c,used,line);
    flag_if(c,used,line);
    for (text,len) in [(a,m),(b,n)]{
        let p=loop_begin(c,i,len,line);
        code(c,text,i,line);c.emit_i32_const(128,line);c.emit_op(Op::I32_GE_U,line);
        c.emit_if(line);c.emit_i32_const(0,line);set(c,used,line);c.emit_end(line);
        loop_end(c,i,p,line);
    }
    c.emit_end(line);
    get(c,used,line);call(c,"wasm:js-number","toI32",1,line);c.emit_if_value(line);
    c.emit_i32_const(128,line);c.emit_op_u16(Op::ARRAY_NEW_DEFAULT,0,line);set(c,used,line);
    let p=loop_begin(c,i,n,line);
    code(c,b,i,line);set(c,j,line);
    get(c,used,line);get(c,j,line); // retained destination operands
    get(c,used,line);get(c,j,line);c.emit_op(Op::ARRAY_GET,line);
    set(c,total,line); // total is dead until counting b completes; no DUP scratch
    get(c,total,line);c.emit_op(Op::REF_IS_NULL,line);c.emit_if_value(line);
    c.emit_f64_const(0.,line);
    c.emit_else(line);get(c,total,line);c.emit_end(line);
    c.emit_f64_const(1.,line);c.emit_op(Op::F64_ADD,line);
    c.emit_op(Op::ARRAY_SET,line);
    loop_end(c,i,p,line);
    c.emit_f64_const(0.,line);set(c,total,line);
    let p=loop_begin(c,i,m,line);
    code(c,a,i,line);set(c,j,line);
    get(c,used,line);get(c,j,line);c.emit_op(Op::ARRAY_GET,line);set(c,n,line); // b length is dead
    get(c,n,line);c.emit_op(Op::REF_IS_NULL,line);c.emit_op(Op::I32_EQZ,line);c.emit_if(line);
    get(c,n,line);c.emit_f64_const(0.,line);c.emit_op(Op::F64_GT,line);c.emit_if(line);
    // Preserve frozen BR1 inside two IFs: consume all positions of this
    // character on its first occurrence in a, not one position per occurrence.
    get(c,used,line);get(c,j,line);c.emit_f64_const(0.,line);
    c.emit_op(Op::ARRAY_SET,line);get(c,total,line);get(c,n,line);c.emit_op(Op::F64_ADD,line);set(c,total,line);
    c.emit_end(line);c.emit_end(line);
    loop_end(c,i,p,line);get(c,total,line);
    c.emit_else(line);
    // Strict lengths and all eight slots already exist; no second reads/spills.
    baseline::emit_prevalidated(chunks,current,line,b,a,used,total,m,n,i,j);
    assert_eq!(chunks[current].local_count,base+6);
    chunks[current].emit_end(line);
}

}
