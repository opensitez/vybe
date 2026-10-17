#![allow(dead_code, unused_imports)]
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

    if enable_myers {
        emit_myers_guard(
            chunk,
            a_slot,
            b_slot,
            m_slot,
            n_slot,
            j_slot,
            cost_slot,
            symmetric_myers,
            line,
        );
        chunk.emit_if_value(line);
        // Keep original argument roots untouched. Only the selected Latin-1
        // branch orients private aliases and dead numeric work slots.
        let (text, pattern) = if symmetric_myers {
            let text = alloc_local(chunk);
            let pattern = alloc_local(chunk);
            lget(chunk, m_slot, line);
            lget(chunk, n_slot, line);
            chunk.emit_op(Op::I32_LT_U, line);
            chunk.emit_if(line);
            lget(chunk, b_slot, line);
            lset(chunk, text, line);
            lget(chunk, a_slot, line);
            lset(chunk, pattern, line);
            lget(chunk, m_slot, line);
            lset(chunk, tmp_slot, line);
            lget(chunk, n_slot, line);
            lset(chunk, m_slot, line);
            lget(chunk, tmp_slot, line);
            lset(chunk, n_slot, line);
            chunk.emit_else(line);
            lget(chunk, a_slot, line);
            lset(chunk, text, line);
            lget(chunk, b_slot, line);
            lset(chunk, pattern, line);
            chunk.emit_end(line);
            (text, pattern)
        } else {
            (a_slot, b_slot)
        };
        emit_myers_body(
            chunk,
            text,
            pattern,
            m_slot,
            n_slot,
            prev_slot,
            curr_slot,
            i_slot,
            j_slot,
            cost_slot,
            tmp_slot,
            v_slot,
            a_char_slot,
            b_chars_slot,
            line,
        );
        chunk.emit_else(line);
    }

    // Strings are immutable. Preserve charAt's UTF-16/code-unit contract,
    // but calculate each b character once rather than in every matrix row.
    // Same private character cache and query order; m<=0 retains an
    // empty buffer. Positive n is already an exact integer length.
    lget(chunk, m_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if(line);
    lget(chunk, n_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_MAX, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, line);
    lset(chunk, b_chars_slot, line);
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
    lget(chunk, j_slot, line);
    lget(chunk, b_slot, line);
    lget(chunk, j_slot, line);
    let idx = chunk.add_import("ecma:string", "charAt");
    chunk.emit_call(idx, 2, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, cache_state, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, line);
    lset(chunk, b_chars_slot, line);
    chunk.emit_end(line);

    // prev[j] = j  (distance from "" to b[..j])
    lget(chunk, n_slot, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_MAX, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, line);
    lset(chunk, prev_slot, line);
    push_const(chunk, Value::F64(0.0), line);
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
    lget(chunk, prev_slot, line);
    lget(chunk, j_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
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
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, curr_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);

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
    chunk.emit_op(Op::F64_LE, line);
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
    // String.equals is an i32 producer. Its negation is already the exact
    // numeric 0/1 substitution cost; F64_ADD accepts that numeric word.
    chunk.emit_op(Op::I32_EQZ, line);
    lset(chunk, cost_slot, line);

    // Retain above before overwriting row[j]. All distance cells are private.
    lget(chunk, prev_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, tmp_slot, line);

    // Retain the write destination under the pure numeric calculation.
    lget(chunk, prev_slot, line);
    lget(chunk, j_slot, line);
    // min(left + 1, above + 1, diagonal + cost)
    lget(chunk, curr_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);

    // ins = prev[j] + 1
    lget(chunk, tmp_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_MIN, line);

    // sub = prev[j-1] + cost
    lget(chunk, v_slot, line);
    lget(chunk, cost_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_MIN, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, curr_slot, line);

    // row[j] = value; retain value as left and the old above as diagonal.
    chunk.emit_op(Op::ARRAY_SET, line);
    lget(chunk, tmp_slot, line);
    lset(chunk, v_slot, line);
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
    if enable_myers {
        chunk.emit_end(line);
    }
    chunk.emit_end(line); // equal-string result / original algorithms
}
// Standard I32 global-Levenshtein Myers scan; b is Latin-1 and 8..=32 units.
// The Ph shift inserts1, retaining the top-row j global-distance boundary.

// Pure core charCodeAt scan; no RegExp/global static state or callbacks.
// j/cost are existing DP locals dead until the fallback initializes them.
fn emit_myers_guard(
    c: &mut Chunk,
    a: u16,
    b: u16,
    m: u16,
    n: u16,
    j: u16,
    flag: u16,
    symmetric: bool,
    line: u32,
) {
    lget(c, n, line);
    c.emit_i32_const(8, line);
    c.emit_op(Op::I32_SUB, line);
    c.emit_i32_const(24, line);
    c.emit_op(Op::I32_LE_U, line);
    if symmetric {
        lget(c, m, line);
        c.emit_i32_const(8, line);
        c.emit_op(Op::I32_SUB, line);
        c.emit_i32_const(24, line);
        c.emit_op(Op::I32_LE_U, line);
        c.emit_op(Op::I32_OR, line);
    }
    c.emit_if_i32(line);
    lget(c, m, line);
    c.emit_i32_const(8, line);
    c.emit_op(Op::I32_GE_S, line);
    if symmetric {
        lget(c, n, line);
        c.emit_i32_const(8, line);
        c.emit_op(Op::I32_GE_S, line);
        c.emit_op(Op::I32_AND, line);
    }
    c.emit_if_i32(line);
    c.emit_i32_const(1, line);
    lset(c, flag, line);
    let charcode = c.add_import("wasm:js-string", "charCodeAt");
    for (text, len) in [(a, m), (b, n)] {
        c.emit_i32_const(0, line);
        lset(c, j, line);
        c.emit_block(line);
        c.emit_loop_s(line);
        lget(c, j, line);
        lget(c, len, line);
        c.emit_op(Op::I32_GE_S, line);
        c.emit_br_if(1, line);
        lget(c, text, line);
        lget(c, j, line);
        c.emit_call(charcode, 2, line);
        c.emit_i32_const(255, line);
        c.emit_op(Op::I32_GT_U, line);
        c.emit_if(line);
        c.emit_i32_const(0, line);
        lset(c, flag, line);
        c.emit_br(2, line);
        c.emit_end(line);
        lget(c, j, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_ADD, line);
        lset(c, j, line);
        c.emit_br(0, line);
        c.emit_end(line);
        c.emit_end(line);
    }
    lget(c, flag, line);
    let to_i32 = c.add_import("wasm:js-number", "toI32");
    c.emit_call(to_i32, 1, line);
    c.emit_else(line);
    c.emit_i32_const(0, line);
    c.emit_end(line);
    c.emit_else(line);
    c.emit_i32_const(0, line);
    c.emit_end(line);
}
fn emit_myers_body(
    c: &mut Chunk,
    a: u16,
    b: u16,
    m: u16,
    n: u16,
    prev: u16,
    curr: u16,
    i: u16,
    old_j: u16,
    cost: u16,
    tmp: u16,
    v: u16,
    a_char: u16,
    b_chars: u16,
    line: u32,
) {
    // Mutually exclusive DP slots have no initialized owner values here.
    // n remains the pattern length until setup/top are complete, then is score.
    // Ph/code share a slot across disjoint setup/main phases. Arguments untouched.
    let top = c.alloc_scratch(1);
    let (mask, j, code, eq, pv, mv, xv, xh, ph, mh, score) = (
        prev, i, a_char, cost, curr, old_j, tmp, v, a_char, b_chars, n,
    );
    let charcode = c.add_import("wasm:js-string", "charCodeAt");
    // Standard GC array.new fills every private mask lane with numeric0.
    // Keep the original first index-slot write/root-release boundary.
    c.emit_i32_const(0, line);
    c.emit_i32_const(256, line);
    c.emit_op_u16(Op::ARRAY_NEW, 0, line);
    lset(c, mask, line);
    c.emit_i32_const(0, line);
    lset(c, j, line);
    // Peq[b[j]] |= 1 << j.
    c.emit_i32_const(0, line);
    lset(c, j, line);
    c.emit_block(line);
    c.emit_loop_s(line);
    lget(c, j, line);
    lget(c, n, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);
    lget(c, b, line);
    lget(c, j, line);
    c.emit_call(charcode, 2, line);
    lset(c, code, line);
    lget(c, mask, line);
    lget(c, code, line);
    lget(c, mask, line);
    lget(c, code, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_i32_const(1, line);
    lget(c, j, line);
    c.emit_op(Op::I32_SHL, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_op(Op::ARRAY_SET, line);
    lget(c, j, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    lset(c, j, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_i32_const(-1, line);
    lset(c, pv, line);
    c.emit_i32_const(0, line);
    lset(c, mv, line);
    c.emit_i32_const(1, line);
    lget(c, n, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_SUB, line);
    c.emit_op(Op::I32_SHL, line);
    lset(c, top, line);
    c.emit_i32_const(0, line);
    lset(c, j, line);
    c.emit_block(line);
    c.emit_loop_s(line);
    lget(c, j, line);
    lget(c, m, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);
    lget(c, mask, line);
    lget(c, a, line);
    lget(c, j, line);
    c.emit_call(charcode, 2, line);
    c.emit_op(Op::ARRAY_GET, line);
    lset(c, eq, line);
    lget(c, eq, line);
    lget(c, mv, line);
    c.emit_op(Op::I32_OR, line);
    lset(c, xv, line);
    lget(c, eq, line);
    lget(c, pv, line);
    c.emit_op(Op::I32_AND, line);
    lget(c, pv, line);
    c.emit_op(Op::I32_ADD, line);
    lget(c, pv, line);
    c.emit_op(Op::I32_XOR, line);
    lget(c, eq, line);
    c.emit_op(Op::I32_OR, line);
    lset(c, xh, line);
    lget(c, mv, line);
    lget(c, xh, line);
    lget(c, pv, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_i32_const(-1, line);
    c.emit_op(Op::I32_XOR, line);
    c.emit_op(Op::I32_OR, line);
    lset(c, ph, line);
    lget(c, pv, line);
    lget(c, xh, line);
    c.emit_op(Op::I32_AND, line);
    lset(c, mh, line);
    for (bits, op) in [(ph, Op::I32_ADD), (mh, Op::I32_SUB)] {
        lget(c, bits, line);
        lget(c, top, line);
        c.emit_op(Op::I32_AND, line);
        c.emit_if(line);
        lget(c, score, line);
        c.emit_i32_const(1, line);
        c.emit_op(op, line);
        lset(c, score, line);
        c.emit_end(line);
    }
    lget(c, ph, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_SHL, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_OR, line);
    lset(c, ph, line);
    lget(c, mh, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_SHL, line);
    lset(c, mh, line);
    lget(c, mh, line);
    lget(c, xv, line);
    lget(c, ph, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_i32_const(-1, line);
    c.emit_op(Op::I32_XOR, line);
    c.emit_op(Op::I32_OR, line);
    lset(c, pv, line);
    lget(c, ph, line);
    lget(c, xv, line);
    c.emit_op(Op::I32_AND, line);
    lset(c, mv, line);
    lget(c, j, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    lset(c, j, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.emit_end(line);
    lget(c, score, line);
    c.emit_op(Op::F64_FROM_I32, line);
}
