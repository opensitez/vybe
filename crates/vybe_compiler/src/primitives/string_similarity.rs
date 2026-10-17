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



fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(value)), line);
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
    let index = code_slot;
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(65, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, index, line);
    chunk.emit_i32_const(25, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_if_i32(line);
    push_str(chunk, "01230120022455012623010202", line);
    lget(chunk, index, line);
    let char_code = chunk.add_import("wasm:js-string", "charCodeAt");
    chunk.emit_call(char_code, 2, line);
    chunk.emit_i32_const(48, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    lset(chunk, digit_slot, line);
}

const SOUNDEX_OWNER: &str = "<primitive-soundex-owned-body>";

/// One module-local Wasm function owns the algorithm; coercion stays at adapters.
pub fn emit_soundex(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    if chunks.len() >= u32::MAX as usize {
        emit_soundex_body(chunks, current, argc, line);
        return;
    }
    let target = if let Some(index) = chunks.iter().position(|c| {
        c.arity == 1 && c.local_names.iter().any(|n|
            n.slot == 0 && n.origin == LocalOrigin::Compiler && n.name == SOUNDEX_OWNER)
    }) {
        index
    } else {
        let mut body = crate::primitives::functions::create_function_chunk("primitive.soundex", 1);
        body.local_count = 1;
        body.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let mut module = vec![body];
        emit_soundex_body(&mut module, 0, 1, 0);
        let mut body = module.remove(0);
        body.emit_op(Op::RETURN, 0);
        body.local_names.push(LocalName::new(0, SOUNDEX_OWNER, LocalOrigin::Compiler));
        let index = chunks.len();
        chunks.push(body);
        index
    };
    if current < target {
        let (callers, callees) = chunks.split_at_mut(target);
        callers[current].emit_direct_defined_call(target as u32, &callees[0], false, line)
            .expect("uncaptured primitive helper");
    } else {
        let (callees, callers) = chunks.split_at_mut(current);
        callers[0].emit_direct_defined_call(target as u32, &callees[target], false, line)
            .expect("uncaptured primitive helper");
    }
}

fn emit_soundex_body(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let last_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let digit_slot = alloc_local(chunk);
    let output_len = alloc_local(chunk);

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
    chunk.emit_op_u16(Op::LOCAL_TEE, n_slot, line);
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
    chunk.emit_i32_const(1, line);
    lset(chunk, output_len, line);

    // last = digit_for(first letter)
    chunk.emit_i32_const(0, line);
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

    // n was cached from the immutable uppercase string before the empty check.
    push_const(chunk, Value::F64(1.0), line);
    lset(chunk, i_slot, line);

    // while i < n && out.length < 4
    let _ = chunk;
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    // compound condition: i < n AND out.length < 4
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_i32(line);
    lget(chunk, output_len, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_LT_S, line);
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
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if(line);
    // digit == "0": reset last
    chunk.emit_i32_const(0, line);
    lset(chunk, last_slot, line);
    chunk.emit_else(line);
    // digit != "0": check if same as last; if not same: append
    lget(chunk, digit_slot, line);
    lget(chunk, last_slot, line);
    chunk.emit_op(Op::I32_NE, line);
    chunk.emit_if(line);
    // different non-zero: append
    lget(chunk, out_slot, line);
    lget(chunk, digit_slot, line);
    chunk.emit_i32_const(48, line);
    chunk.emit_op(Op::I32_ADD, line);
    let from_code = chunk.add_import("wasm:js-string", "fromCharCode");
    chunk.emit_call(from_code, 1, line);
    crate::primitives::strings::emit_concat(chunk, 2, line);
    lset(chunk, out_slot, line);
    lget(chunk, output_len, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, output_len, line);
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

    // The result starts as one UTF-16 unit and only gains ASCII digits.
    lget(chunk, out_slot, line);
    chunk.emit_i32_const(4, line);
    push_str(chunk, "0", line);
    let pad_end = chunk.add_import("ecma:string", "padEnd");
    chunk.emit_call(pad_end, 3, line);
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
    // Original DP needs11 slots; preserve exact emitter at the old frame limit.
    let enable_myers = chunks[current].local_count <= u16::MAX - 11 - 1;
    let symmetric_myers = chunks[current].local_count <= u16::MAX - 11 - 3;
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

    // Both strict length reads precede equality. Identical strings have zero
    // distance, including the VM's replacement-unit representation. Keep the
    // original fallback for signed-wrapped lengths; equal strings have equal
    // standard String.length results.
    lget(chunk, m_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GE_S, line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    crate::primitives::strings::emit_str_equals(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_else(line);

    // With n>=0, m<=0 retains the old n result; n==0 retains max(m,0).
    // Negative n keeps the original array/bounds behavior. This also makes
    // zero-bound character-cache guard redundant in the remaining DP path.
    lget(chunk, m_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_LE_S, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_OR, line);
    lget(chunk, n_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    lget(chunk, m_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_MAX, line);
    chunk.emit_else(line);

    if enable_myers {
        emit_myers_guard(chunk, a_slot, b_slot, m_slot, n_slot, j_slot, cost_slot, symmetric_myers, line);
        chunk.emit_if_value(line);
        // Keep original argument roots untouched. Only the selected Latin-1
        // branch orients private aliases and dead numeric work slots.
        let (text, pattern) = if symmetric_myers {
            let text = alloc_local(chunk);
            let pattern = alloc_local(chunk);
            lget(chunk, m_slot, line); lget(chunk, n_slot, line);
            chunk.emit_op(Op::I32_LT_U, line);
            chunk.emit_if(line);
            lget(chunk, b_slot, line); lset(chunk, text, line);
            lget(chunk, a_slot, line); lset(chunk, pattern, line);
            lget(chunk, m_slot, line); lset(chunk, tmp_slot, line);
            lget(chunk, n_slot, line); lset(chunk, m_slot, line);
            lget(chunk, tmp_slot, line); lset(chunk, n_slot, line);
            chunk.emit_else(line);
            lget(chunk, a_slot, line); lset(chunk, text, line);
            lget(chunk, b_slot, line); lset(chunk, pattern, line);
            chunk.emit_end(line);
            (text, pattern)
        } else { (a_slot, b_slot) };
        emit_myers_body(chunk, text, pattern, m_slot, n_slot, prev_slot, curr_slot, i_slot, j_slot, cost_slot, tmp_slot, v_slot, a_char_slot, b_chars_slot, line);
        chunk.emit_else(line);
    }

    // Strings are immutable. Preserve charAt's UTF-16/code-unit contract,
    // but calculate each b character once rather than in every matrix row.
    // Same private character cache and query order; m<=0 retains an
    // empty buffer. Positive n is already an exact integer length.
    lget(chunk,n_slot,line);chunk.emit_i32_const(0,line);chunk.emit_op(Op::F64_MAX,line);
    chunk.emit_op_u16(Op::ARRAY_NEW_DEFAULT,0,line);lset(chunk,b_chars_slot,line);
    chunk.emit_i32_const(0,line);
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
    lget(chunk, j_slot, line);
    lget(chunk, b_slot, line);
    lget(chunk, j_slot, line);
    let idx = chunk.add_import("ecma:string", "charAt");
    chunk.emit_call(idx, 2, line);
    chunk.emit_op(Op::ARRAY_SET,line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, cache_state, line);
    let chunk=&mut chunks[current];

    // prev[j] = j  (distance from "" to b[..j])
    lget(chunk,n_slot,line);chunk.emit_f64_const(1.0,line);chunk.emit_op(Op::F64_ADD,line);
    chunk.emit_i32_const(0,line);chunk.emit_op(Op::F64_MAX,line);
    chunk.emit_op_u16(Op::ARRAY_NEW_DEFAULT,0,line);lset(chunk,prev_slot,line);
    chunk.emit_i32_const(0,line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    let init_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_LE, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk,prev_slot,line);lget(chunk,j_slot,line);lget(chunk,j_slot,line);
    chunk.emit_op(Op::ARRAY_SET,line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, init_state, line);

    // Outer loop: for i in 1..=m
    push_const(&mut chunks[current], Value::F64(1.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let outer_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, m_slot, line);
    chunk.emit_op(Op::F64_LE, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    // One in-place distance row. curr_slot is the current row's left value;
    // v_slot retains the previous row's diagonal, initially i-1.
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, v_slot, line);
    lget(chunk, prev_slot, line);
    chunk.emit_i32_const(0,line);
    lget(chunk, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, curr_slot, line);
    chunk.emit_op(Op::ARRAY_SET,line);

    let chunk = &mut chunks[current];
    lget(chunk, n_slot, line);chunk.emit_i32_const(0,line);chunk.emit_op(Op::I32_GT_S,line);chunk.emit_if(line);
    lget(chunk, a_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, a_char_slot, line);chunk.emit_end(line);

    push_const(&mut chunks[current], Value::F64(1.0), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
    let inner_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_LE, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    // Keep the destination live. Form diagonal+cost before above replaces
    // the diagonal. Finite nonnegative integer candidates make min associative.
    lget(chunk,prev_slot,line);lget(chunk,j_slot,line);lget(chunk,v_slot,line);
    lget(chunk,a_char_slot,line);lget(chunk,b_chars_slot,line);lget(chunk,j_slot,line);
    push_const(chunk,Value::F64(1.0),line);chunk.emit_op(Op::F64_SUB,line);
    chunk.emit_op(Op::ARRAY_GET,line);crate::primitives::strings::emit_str_equals(chunk,line);
    chunk.emit_op(Op::I32_EQZ,line);chunk.emit_op(Op::F64_ADD,line);
    lget(chunk,prev_slot,line);lget(chunk,j_slot,line);chunk.emit_op(Op::ARRAY_GET,line);
    chunk.emit_op_u16(Op::LOCAL_TEE,v_slot,line);
    push_const(chunk,Value::F64(1.0),line);chunk.emit_op(Op::F64_ADD,line);
    chunk.emit_op(Op::F64_MIN,line);
    lget(chunk,curr_slot,line);push_const(chunk,Value::F64(1.0),line);
    chunk.emit_op(Op::F64_ADD,line);chunk.emit_op(Op::F64_MIN,line);
    chunk.emit_op_u16(Op::LOCAL_TEE,curr_slot,line);chunk.emit_op(Op::ARRAY_SET,line);
    let chunk = &mut chunks[current];

    // j++
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, inner_state, line);
    let chunk = &mut chunks[current];

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
    if enable_myers { chunk.emit_end(line); }
    chunk.emit_end(line); // zero-bound result / original algorithms
    chunk.emit_end(line); // equal-string result / original algorithms
}
// Standard I32 global-Levenshtein Myers scan; b is Latin-1 and 8..=32 units.
// The Ph shift inserts1, retaining the top-row j global-distance boundary.

// Pure core charCodeAt scan; no RegExp/global static state or callbacks.
// j/cost are existing DP locals dead until the fallback initializes them.
fn emit_myers_guard(c:&mut Chunk,a:u16,b:u16,m:u16,n:u16,j:u16,flag:u16,symmetric:bool,line:u32){
    lget(c,n,line);c.emit_i32_const(8,line);c.emit_op(Op::I32_SUB,line);
    c.emit_i32_const(24,line);c.emit_op(Op::I32_LE_U,line);
    if symmetric {
        lget(c,m,line);c.emit_i32_const(8,line);c.emit_op(Op::I32_SUB,line);
        c.emit_i32_const(24,line);c.emit_op(Op::I32_LE_U,line);
        c.emit_op(Op::I32_OR,line);
    }
    c.emit_if_i32(line);
    lget(c,m,line);c.emit_i32_const(8,line);c.emit_op(Op::I32_GE_S,line);
    if symmetric {
        lget(c,n,line);c.emit_i32_const(8,line);c.emit_op(Op::I32_GE_S,line);
        c.emit_op(Op::I32_AND,line);
    }
    c.emit_if_i32(line);
    c.emit_i32_const(1,line);lset(c,flag,line);
    let charcode=c.add_import("wasm:js-string","charCodeAt");
    for (text,len) in [(a,m),(b,n)] {
        c.emit_i32_const(0,line);lset(c,j,line);
        c.emit_block(line);c.emit_loop_s(line);
        lget(c,j,line);lget(c,len,line);c.emit_op(Op::I32_GE_S,line);c.emit_br_if(1,line);
        lget(c,text,line);lget(c,j,line);c.emit_call(charcode,2,line);
        c.emit_i32_const(255,line);c.emit_op(Op::I32_GT_U,line);c.emit_if(line);
        c.emit_i32_const(0,line);lset(c,flag,line);c.emit_br(2,line);
        c.emit_end(line);
        lget(c,j,line);c.emit_i32_const(1,line);c.emit_op(Op::I32_ADD,line);lset(c,j,line);
        c.emit_br(0,line);c.emit_end(line);c.emit_end(line);
    }
    lget(c,flag,line);let to_i32=c.add_import("wasm:js-number","toI32");c.emit_call(to_i32,1,line);
    c.emit_else(line);c.emit_i32_const(0,line);c.emit_end(line);
    c.emit_else(line);c.emit_i32_const(0,line);c.emit_end(line);
}
fn emit_myers_body(c:&mut Chunk,a:u16,b:u16,m:u16,n:u16,prev:u16,curr:u16,i:u16,old_j:u16,cost:u16,tmp:u16,v:u16,a_char:u16,b_chars:u16,line:u32){
    // Mutually exclusive DP slots have no initialized owner values here.
    // n remains the pattern length until setup/top are complete, then is score.
    // Ph/code share a slot across disjoint setup/main phases. Arguments untouched.
    let top=c.alloc_scratch(1);
    let (mask,j,code,eq,pv,mv,xv,xh,ph,mh,score)=(prev,i,a_char,cost,curr,old_j,tmp,v,a_char,b_chars,n);
    let charcode=c.add_import("wasm:js-string","charCodeAt");
    // Standard GC array.new fills every private mask lane with numeric0.
    // Keep the original first index-slot write/root-release boundary.
    c.emit_i32_const(0,line);c.emit_i32_const(256,line);
    c.emit_op_u16(Op::ARRAY_NEW,0,line);lset(c,mask,line);
    c.emit_i32_const(0,line);lset(c,j,line);
    // Peq[b[j]] |= 1 << j.
    c.emit_i32_const(0,line);lset(c,j,line);c.emit_block(line);c.emit_loop_s(line);
    lget(c,j,line);lget(c,n,line);c.emit_op(Op::I32_GE_S,line);c.emit_br_if(1,line);
    lget(c,b,line);lget(c,j,line);c.emit_call(charcode,2,line);lset(c,code,line);
    lget(c,mask,line);lget(c,code,line);
    lget(c,mask,line);lget(c,code,line);c.emit_op(Op::ARRAY_GET,line);
    c.emit_i32_const(1,line);lget(c,j,line);c.emit_op(Op::I32_SHL,line);c.emit_op(Op::I32_OR,line);c.emit_op(Op::ARRAY_SET,line);
    lget(c,j,line);c.emit_i32_const(1,line);c.emit_op(Op::I32_ADD,line);lset(c,j,line);c.emit_br(0,line);c.emit_end(line);c.emit_end(line);
    c.emit_i32_const(-1,line);lset(c,pv,line);c.emit_i32_const(0,line);lset(c,mv,line);
    c.emit_i32_const(1,line);lget(c,n,line);c.emit_i32_const(1,line);c.emit_op(Op::I32_SUB,line);c.emit_op(Op::I32_SHL,line);lset(c,top,line);
    c.emit_i32_const(0,line);lset(c,j,line);c.emit_block(line);c.emit_loop_s(line);
    lget(c,j,line);lget(c,m,line);c.emit_op(Op::I32_GE_S,line);c.emit_br_if(1,line);
    lget(c,mask,line);lget(c,a,line);lget(c,j,line);c.emit_call(charcode,2,line);c.emit_op(Op::ARRAY_GET,line);lset(c,eq,line);
    lget(c,eq,line);lget(c,mv,line);c.emit_op(Op::I32_OR,line);lset(c,xv,line);
    lget(c,eq,line);lget(c,pv,line);c.emit_op(Op::I32_AND,line);lget(c,pv,line);c.emit_op(Op::I32_ADD,line);
    lget(c,pv,line);c.emit_op(Op::I32_XOR,line);lget(c,eq,line);c.emit_op(Op::I32_OR,line);lset(c,xh,line);
    lget(c,mv,line);lget(c,xh,line);lget(c,pv,line);c.emit_op(Op::I32_OR,line);c.emit_i32_const(-1,line);c.emit_op(Op::I32_XOR,line);c.emit_op(Op::I32_OR,line);lset(c,ph,line);
    lget(c,pv,line);lget(c,xh,line);c.emit_op(Op::I32_AND,line);lset(c,mh,line);
    for (bits,op) in [(ph,Op::I32_ADD),(mh,Op::I32_SUB)]{
        lget(c,bits,line);lget(c,top,line);c.emit_op(Op::I32_AND,line);c.emit_if(line);
        lget(c,score,line);c.emit_i32_const(1,line);c.emit_op(op,line);lset(c,score,line);c.emit_end(line);
    }
    lget(c,ph,line);c.emit_i32_const(1,line);c.emit_op(Op::I32_SHL,line);c.emit_i32_const(1,line);c.emit_op(Op::I32_OR,line);lset(c,ph,line);
    lget(c,mh,line);c.emit_i32_const(1,line);c.emit_op(Op::I32_SHL,line);lset(c,mh,line);
    lget(c,mh,line);lget(c,xv,line);lget(c,ph,line);c.emit_op(Op::I32_OR,line);c.emit_i32_const(-1,line);c.emit_op(Op::I32_XOR,line);c.emit_op(Op::I32_OR,line);lset(c,pv,line);
    lget(c,ph,line);lget(c,xv,line);c.emit_op(Op::I32_AND,line);lset(c,mv,line);
    lget(c,j,line);c.emit_i32_const(1,line);c.emit_op(Op::I32_ADD,line);lset(c,j,line);c.emit_br(0,line);c.emit_end(line);c.emit_end(line);
    lget(c,score,line);c.emit_op(Op::F64_FROM_I32,line);
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

    // charCodeAt produces an i32 UTF-16 unit. Guard the A..Z index before
    // shifting so Wasm's modulo-32 shift cannot accept punctuation/non-ASCII.
    // Reuse the former flag slot; no caller roots or frame size change.
    let index_slot = alloc_local(chunk);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(65, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, index_slot, line);
    chunk.emit_i32_const(25, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(45071982, line);
    lget(chunk, index_slot, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    crate::primitives::strings::emit_concat(chunk, 2, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line); // end consonant mask
    chunk.emit_end(line); // end ASCII range
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
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};


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
    // A true used lane accompanies each increment. Once all B lanes are
    // consumed, no remaining original scan can read/compare characters or
    // change total. Exit both loops only on this proven match path; leave
    // disjoint/unused-character paths completely unchanged.
    lget(chunk, total_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_GE, line);
    // Two IFs and the inner BLOCK/LOOP are inside the outer BLOCK/LOOP.
    chunk.emit_br_if(outer_state.break_depth(4) as u32, line);
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
    // Strict String.length produces I32. Private indices never exceed the
    // signed-positive bound; negative wrapped lengths execute zero iterations.
    c.emit_i32_const(0,l);set(c,index,l);
    let block=c.emit_block(l);let (lp,_)=c.emit_loop_s(l);
    get(c,index,l);get(c,len,l);c.emit_op(Op::I32_LT_S,l);c.emit_op(Op::I32_EQZ,l);c.emit_br_if(1,l);
    (block,lp)
}
fn loop_end(c:&mut Chunk,index:u16,patch:(usize,usize),l:u32){
    get(c,index,l);c.emit_i32_const(1,l);c.emit_op(Op::I32_ADD,l);set(c,index,l);
    c.emit_br(0,l);c.emit_end(l);c.patch_loop(patch.1);c.emit_end(l);c.patch_block(patch.0);
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
        c.emit_if(line);c.emit_i32_const(0,line);set(c,used,line);
        // All further strict indexed reads are pure and in bounds. Once a
        // non-ASCII unit is found, both remaining scans cannot change used.
        // Depth3 leaves this IF, loop, block and enclosing classification IF.
        c.emit_br(3,line);c.emit_end(line);
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
    // Unicode-heavy fallback only. Keep both strict length reads and the
    // existing ASCII classification/frequency path in their original order.
    // Bitmap is a private standard ref GC array; Null default means zero word.
    let c = &mut chunks[current];
    get(c,m,line); get(c,n,line); c.emit_op(Op::F64_MUL,line);
    c.emit_f64_const(16384.,line); c.emit_op(Op::F64_GE,line);
    get(c,m,line); c.emit_i32_const(0,line); c.emit_op(Op::I32_GT_S,line);
    c.emit_op(Op::I32_AND,line);
    get(c,n,line); c.emit_i32_const(0,line); c.emit_op(Op::I32_GT_S,line);
    c.emit_op(Op::I32_AND,line);
    c.emit_if_value(line);
    // Strict lengths precede this pure equality; positive n retains F64 ABI.
    get(c,a,line); get(c,b,line); call(c,"wasm:js-string","equals",2,line);
    c.emit_if_value(line);
    get(c,n,line); c.emit_op(Op::F64_FROM_I32,line);
    c.emit_else(line);
    emit_unit_membership(c,a,b,used,total,m,n,i,j,line);
    c.emit_end(line);
    c.emit_else(line);
    baseline::emit_prevalidated(chunks,current,line,b,a,used,total,m,n,i,j);
    chunks[current].emit_end(line);
    assert_eq!(chunks[current].local_count,base+6);
    chunks[current].emit_end(line);
}


/// Preserve ecma:charAt's existing UTF-16 unit/replacement policy before
/// extracting the resulting unit. Native charAt keeps lone surrogates; the
/// VM replaces them with U+FFFD. Both remain their established equality keys.
fn legacy_unit(c:&mut Chunk,text:u16,index:u16,l:u32) {
    get(c,text,l); get(c,index,l); call(c,"ecma:string","charAt",2,l);
    c.emit_i32_const(0,l); call(c,"wasm:js-string","charCodeAt",2,l);
}
fn word_index(c:&mut Chunk,key:u16,l:u32) {
    get(c,key,l); c.emit_i32_const(5,l); c.emit_op(Op::I32_SHR_U,l);
}
fn unit_bit(c:&mut Chunk,key:u16,l:u32) {
    c.emit_i32_const(1,l); get(c,key,l); c.emit_i32_const(31,l);
    c.emit_op(Op::I32_AND,l); c.emit_op(Op::I32_SHL,l);
}
/// Existing legacy count is the number of B units occurring at least once in
/// A, not the PHP recursive similarity score. This computes that exact count.
/// total is a dead word scratch until the A pass ends; m is dead thereafter.
/// No extra local or frame allocation; the original fallback stays unchanged.
fn emit_unit_membership(c:&mut Chunk,a:u16,b:u16,used:u16,total:u16,
    m:u16,n:u16,i:u16,j:u16,l:u32) {
    c.emit_i32_const(2048,l); c.emit_op_u16(Op::ARRAY_NEW_DEFAULT,0,l);
    set(c,used,l);
    let p=loop_begin(c,i,m,l);
    legacy_unit(c,a,i,l); set(c,j,l);
    get(c,used,l); word_index(c,j,l); // retained array/index destination
    get(c,used,l); word_index(c,j,l); c.emit_op(Op::ARRAY_GET,l);
    set(c,total,l);
    // Void branch updates only this dead private word scratch. Keeping the
    // checked numeric cast supplies the OR operand without I32 block metadata;
    // the existing single-site primitive inliner can preserve all metadata.
    get(c,total,l); c.emit_op(Op::REF_IS_NULL,l); c.emit_if(l);
    c.emit_i32_const(0,l); set(c,total,l);
    c.emit_end(l);
    get(c,total,l); call(c,"wasm:js-number","toI32",1,l);
    unit_bit(c,j,l); c.emit_op(Op::I32_OR,l); c.emit_op(Op::ARRAY_SET,l);
    loop_end(c,i,p,l);
    c.emit_f64_const(0.,l); set(c,total,l);
    let p=loop_begin(c,i,n,l);
    legacy_unit(c,b,i,l); set(c,j,l);
    get(c,used,l); word_index(c,j,l); c.emit_op(Op::ARRAY_GET,l);
    set(c,m,l); // A length was consumed; input A itself stays retained.
    get(c,m,l); c.emit_op(Op::REF_IS_NULL,l); c.emit_op(Op::I32_EQZ,l);
    c.emit_if(l);
    get(c,m,l); call(c,"wasm:js-number","toI32",1,l);
    unit_bit(c,j,l); c.emit_op(Op::I32_AND,l); c.emit_if(l);
    increment(c,total,l);
    c.emit_end(l); c.emit_end(l);
    loop_end(c,i,p,l); get(c,total,l);
}

}
