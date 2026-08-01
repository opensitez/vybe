//! `string_encoding` — representation-change adapter primitives.
//!
//! Text to another representation and back: base64, hex, rot13,
//! quoted-printable, uuencode. **URL percent-encoding is NOT here** — a URL is
//! a structured format and lives in `primitives::url` with the rest of its
//! domain (splitting, joining, query strings). These ROUND-TRIP, which is what separates
//! them from `string_escaping` (making text safe to embed in a host syntax,
//! often one-way).
//!
//! ECMA-262 offers none of them, so every language that exposes them has had to
//! write its own — PHP in `string_adapter.rs`, Python in a whole separate
//! `url_adapter.rs`. Percent-encoding was implemented twice on the platform
//! before this module existed.
//!
//! **No coercion here.** Every function takes STRINGS on the stack; how a
//! language turns a non-string into a string is its own rule and stays at its
//! call site (PHP coerces silently and differently from ECMA; Python raises).

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

fn emit_floordiv(chunk: &mut Chunk, slot: u16, div: f64, line: u32) {
    lget(chunk, slot, line);
    push_const(chunk, Value::F64(div), line);
    chunk.emit_op(Op::F64_DIV, line);
    crate::primitives::math::emit_floor(chunk, line);
}

fn emit_hexval(chunk: &mut Chunk, code_slot: u16, line: u32) {
    // 48-57 → -48 ; 65-70 → -55 ; 97-102 → -87 (upper/lower A-F)
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(57.0), line);
    crate::primitives::ops::emit_dyn_le(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(48.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_else(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(97.0), line);
    crate::primitives::ops::emit_dyn_ge(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(87.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_else(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(55.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
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
    crate::primitives::ops::emit_dyn_ge(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    push_const(chunk, Value::F64(96.0), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(32.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_end(line);
}

fn emit_uu_enc(chunk: &mut Chunk, c_slot: u16, line: u32) {
    lget(chunk, c_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
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

pub fn emit_base64_decode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let atob = chunks[0].add_import("ecma:string", "atob");
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
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::I32_LT_S, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
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
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, valid_slot, line);
    chunk.emit_if_value(line);
    lget(chunk, s_slot, line);
    chunk.emit_call(atob, 1, line);
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    lget(chunk, s_slot, line);
    chunk.emit_call(atob, 1, line);
    chunk.emit_end(line);
}

pub fn emit_bin2hex(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let hi_slot = alloc_local(chunk);
    let lo_slot = alloc_local(chunk);

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
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    // hi = (code >> 4) & 0xF; lo = code & 0xF
    let table = "0123456789abcdef";
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    lset(chunk, hi_slot, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    crate::primitives::math::emit_c_fmod(chunk, line);
    lset(chunk, lo_slot, line);

    // out += table.charAt(hi) + table.charAt(lo)
    lget(chunk, out_slot, line);
    push_str(chunk, table, line);
    lget(chunk, hi_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    crate::primitives::ops::emit_dyn_add(chunk, line);
    push_str(chunk, table, line);
    lget(chunk, lo_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
}

pub fn emit_hex2bin(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let hi_slot = alloc_local(chunk);
    let lo_slot = alloc_local(chunk);

    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
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
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    // condition: i + 1 < len
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, len_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    let table = "0123456789abcdef";
    push_str(chunk, table, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, hi_slot, line);
    push_str(chunk, table, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, lo_slot, line);

    // if hi<0 || lo<0: return false
    lget(chunk, hi_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_end(line);
    lget(chunk, lo_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_end(line);

    // out += String.fromCharCode((hi << 4) | lo)
    lget(chunk, out_slot, line);
    lget(chunk, hi_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    lget(chunk, lo_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);

    // i += 2
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
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
    push_str(chunk, "", line);
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
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
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
    crate::primitives::ops::emit_dyn_ge(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    //   if code <= 90 → uppercase A-Z
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(90.0), line);
    crate::primitives::ops::emit_dyn_le(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    emit_rot13_range(chunk, code_slot, 65.0, tmp_slot, line);
    lset(chunk, rot_slot, line);
    chunk.emit_else(line); // code > 90 (still inside code >= 65 branch)
    //   else if code >= 97 → check lowercase a-z
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(97.0), line);
    crate::primitives::ops::emit_dyn_ge(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    //     if code <= 122 → lowercase a-z
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(122.0), line);
    crate::primitives::ops::emit_dyn_le(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
}

pub fn emit_quoted_printable_encode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let hi_slot = alloc_local(chunk);
    let hex_slot = alloc_local(chunk);

    lset(chunk, s_slot, line);
    push_str(chunk, "0123456789ABCDEF", line);
    lset(chunk, hex_slot, line);
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
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    // need = code==61 | code>126 | (code<32 & code!=9)
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(61.0), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(126.0), line);
    crate::primitives::ops::emit_dyn_gt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(32.0), line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(9.0), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    // out += "=" + hex[hi] + hex[lo]
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    crate::primitives::math::emit_floor(chunk, line);
    lset(chunk, hi_slot, line);
    lget(chunk, out_slot, line);
    push_str(chunk, "=", line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lget(chunk, hex_slot, line);
    lget(chunk, hi_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lget(chunk, hex_slot, line);
    // lo = code - hi*16
    lget(chunk, code_slot, line);
    lget(chunk, hi_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
}

pub fn emit_quoted_printable_decode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let c1_slot = alloc_local(chunk);
    let c2_slot = alloc_local(chunk);
    let byte_slot = alloc_local(chunk);

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
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    // if code == '=' (61)
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(61.0), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // c1 = charCodeAt(i+1)
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, c1_slot, line);
    // soft break? c1==13 | c1==10
    lget(chunk, c1_slot, line);
    push_const(chunk, Value::F64(13.0), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, c1_slot, line);
    push_const(chunk, Value::F64(10.0), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    // soft break: skip '=' + CR (+ LF if CRLF). i += (c1==13 ? 2 : 1) ... plus the '='
    // Advance past '=' and the CR/LF; if CRLF advance one more.
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    // if c1==13: also skip the following LF
    lget(chunk, c1_slot, line);
    push_const(chunk, Value::F64(13.0), line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    // =XX hex escape
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, c2_slot, line);
    emit_hexval(chunk, c1_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_hexval(chunk, c2_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, byte_slot, line);
    lget(chunk, out_slot, line);
    lget(chunk, byte_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(3.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    // literal char
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
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

    lset(chunk, s_slot, line);
    lget(chunk, s_slot, line);
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
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
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
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    emit_floordiv(chunk, b0_slot, 4.0, line);
    lset(chunk, c_slot, line);
    lget(chunk, out_slot, line);
    emit_uu_enc(chunk, c_slot, line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    // c1 = (b0 mod 4)*16 + floor(b1/16)
    emit_modn(chunk, b0_slot, 4.0, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_floordiv(chunk, b1_slot, 16.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, c_slot, line);
    lget(chunk, out_slot, line);
    emit_uu_enc(chunk, c_slot, line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    // c2 = (b1 mod 16)*4 + floor(b2/64)
    emit_modn(chunk, b1_slot, 16.0, line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_floordiv(chunk, b2_slot, 64.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, c_slot, line);
    lget(chunk, out_slot, line);
    emit_uu_enc(chunk, c_slot, line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    // c3 = b2 mod 64
    emit_modn(chunk, b2_slot, 64.0, line);
    lset(chunk, c_slot, line);
    lget(chunk, out_slot, line);
    emit_uu_enc(chunk, c_slot, line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(3.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    push_str(chunk, "\n", line);
    crate::primitives::ops::emit_dyn_add(chunk, line);
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

    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    // count = charCodeAt(0) - 32
    lget(chunk, s_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::F64(32.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, count_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    lset(chunk, pos_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, done_slot, line);

    let _ = chunk;
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, done_slot, line);
    lget(chunk, count_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond(chunks, current, line);
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
    emit_floordiv(chunk, c1, 16.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, byte_slot, line);
    lget(chunk, out_slot, line);
    lget(chunk, byte_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, done_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, done_slot, line);

    // if done < count: emit b1
    lget(chunk, done_slot, line);
    lget(chunk, count_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // b1 = (c1 mod 16)*16 + floor(c2/4)
    emit_modn(chunk, c1, 16.0, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_floordiv(chunk, c2, 4.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, byte_slot, line);
    lget(chunk, out_slot, line);
    lget(chunk, byte_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "fromCharCode");
        chunk.emit_call(idx, 1, line);
    }
    crate::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, done_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, done_slot, line);
    chunk.emit_end(line);

    // if done < count: emit b2
    lget(chunk, done_slot, line);
    lget(chunk, count_slot, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // b2 = (c2 mod 4)*64 + c3
    emit_modn(chunk, c2, 4.0, line);
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
    crate::primitives::ops::emit_dyn_add(chunk, line);
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
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
}
