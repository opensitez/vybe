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

// All output fragments below are proven strings (literals, charAt, or
// fromCharCode). Collect them privately and join once, avoiding quadratic
// copying and dynamic addition/coercion dispatch on every append.
fn emit_parts_new(chunk: &mut Chunk, line: u32) {
    let mut imports = Chunk::new("imports");
    crate::primitives::collections::emit_array_new_into(&mut imports, chunk, 0, line);
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

#[allow(dead_code)]
fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, argc, line);
}

fn emit_base64_code_is_allowed(chunk: &mut Chunk, code_slot: u16, line: u32) {
    crate::primitives::base64::emit_code_is_allowed(chunk, code_slot, false, line);
}

fn emit_base64_lenient_filter(chunks: &mut [Chunk], current: usize, s_slot: u16, line: u32) {
    lget(&mut chunks[current], s_slot, line);
    crate::primitives::base64::emit_filter_lenient(chunks, current, line);
}

fn emit_uu_div_pow2(chunk: &mut Chunk, slot: u16, div: f64, line: u32) {
    // UU-only: callers supply bounded nonnegative UTF16 units or masked sextets.
    // Divisors are powers of two (4, 16, 64); no byte narrowing is performed.
    lget(chunk, slot, line);
    chunk.emit_i32_const((div as u32).trailing_zeros() as i32, line);
    chunk.emit_op(Op::I32_SHR_U, line);
}

fn emit_hexval(chunk: &mut Chunk, code_slot: u16, line: u32) {
    // Both callers supply bounded charCodeAt results. Unsigned subtraction
    // rejects units below each range, including the wrapping negative result.
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(48, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_i32_const(9, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_if_value(line);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(48, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_else(line);
    // Fold ASCII case only for recognition; non-ASCII units cannot enter.
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_i32_const(97, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_i32_const(5, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_if_value(line);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_i32_const(87, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_uu_mod_pow2(chunk: &mut Chunk, slot: u16, modulus: f64, line: u32) {
    // The same bounded UU integer domain proves this mask equals remainder.
    lget(chunk, slot, line);
    chunk.emit_i32_const(modulus as i32 - 1, line);
    chunk.emit_op(Op::I32_AND, line);
}

fn emit_uu_dec(chunk: &mut Chunk, line: u32) {
    // charCodeAt returns 0..65535. Wrapping subtraction and the low-six-bit
    // mask preserve the original signed subtraction followed by &63.
    chunk.emit_i32_const(32, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_i32_const(63, line);
    chunk.emit_op(Op::I32_AND, line);
}

fn emit_uu_enc(chunk: &mut Chunk, c_slot: u16, line: u32) {
    // Each packed integer is bounded by 16383 (UTF16 units are not bytes).
    // Standard select implements c ? c+32 : 96 without a branch or narrowing.
    lget(chunk, c_slot, line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_i32_const(96, line);
    lget(chunk, c_slot, line);
    chunk.emit_op(Op::SELECT, line);
    let idx = chunk.add_import("wasm:js-string", "fromCharCode");
    chunk.emit_call(idx, 1, line);
}

pub fn emit_base64_decode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        let (name, arity) = if argc < 2 {
            ("__vybe_encoding_base64_decode_lenient", 1)
        } else {
            ("__vybe_encoding_base64_decode", 2)
        };
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current],
            name,
            arity,
            line,
        );
        return;
    }
    // Absence of a strict argument proves the lenient default. Do not emit
    // strict validation or a dynamic truthiness ladder for a known default.
    if argc < 2 {
        let input = alloc_local(&mut chunks[current]);
        lset(&mut chunks[current], input, line);
        emit_base64_lenient_filter(chunks, current, input, line);
        crate::primitives::base64::emit_decode_binary_string(chunks, current, line);
        return;
    }

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
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
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
    // All remaining reads would only recheck this immutable string. The
    // failure result is fixed; leave this private scan at its first invalid unit.
    chunk.emit_br(loop_state.break_depth(1) as u32, line);
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
    let _ = chunk;
    crate::primitives::base64::emit_decode_binary_string(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    let _ = chunk;
    emit_base64_lenient_filter(chunks, current, s_slot, line);
    let chunk = &mut chunks[current];
    let _ = chunk;
    crate::primitives::base64::emit_decode_binary_string(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
}

pub fn emit_bin2hex(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current],
            "__vybe_encoding_bin2hex",
            1,
            line,
        );
        return;
    }

    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let hi_slot = alloc_local(chunk);
    let lo_slot = alloc_local(chunk);

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
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
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
    // charCodeAt yields exactly one unsigned UTF16 unit. Preserve the
    // legacy hi>15 empty-table result for non-Latin1 units: do not mask hi.
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    lset(chunk, hi_slot, line);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(15, line);
    chunk.emit_op(Op::I32_AND, line);
    lset(chunk, lo_slot, line);

    // out += table.charAt(hi) + table.charAt(lo)
    lget(chunk, out_slot, line);
    push_str(chunk, table, line);
    lget(chunk, hi_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    emit_append_part(chunk, out_slot, line);
    push_str(chunk, table, line);
    lget(chunk, lo_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    emit_append_part(chunk, out_slot, line);
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
    emit_parts_join(chunk, line);
}

fn hex_import(c: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = c.add_import(module, name);
    c.emit_call(idx, argc, line);
}
pub fn emit_hex2bin(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current],
            "__vybe_encoding_hex2bin",
            1,
            line,
        );
        return;
    }

    emit_hex_decode_body(chunks, current, line, false);
}

/// Decode to a fresh plain byte array, preserving the composed decoder's
/// invalid-input trap. Language-specific null/error policy stays in adapters.
pub fn emit_hex2bytes(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current], "__vybe_encoding_hex2bytes", 1, line,
        );
        return;
    }
    emit_hex_decode_body(chunks, current, line, true);
}

fn emit_hex_decode_body(chunks: &mut [Chunk], current: usize, line: u32, array: bool) {
    // Canonical roundtrip rejects Buffer.from's accepted partial prefixes.
    let c = &mut chunks[current];
    let src = c.alloc_scratch(2);
    let bytes = src + 1;
    hex_import(c, "ecma:string", "toLowerCase", 1, line);
    c.emit_op_u16(Op::LOCAL_SET, src, line);
    lget(c, src, line);
    hex_import(c, "wasm:js-string", "length", 1, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if_value(line);
    if array {
        c.emit_i32_const(0, line);
        hex_import(c, "vybe:js-array", "newWithLength", 1, line);
    } else {
        c.emit_string_const("", line);
    }
    c.emit_else(line);
    lget(c, src, line);
    c.emit_string_const("hex", line);
    hex_import(c, "node:buffer", "from", 2, line);
    c.emit_op_u16(Op::LOCAL_SET, bytes, line);
    lget(c, src, line);
    lget(c, bytes, line);
    c.emit_string_const("hex", line);
    hex_import(c, "node:buffer", "toString", 2, line);
    hex_import(c, "wasm:js-string", "equals", 2, line);
    c.emit_if_value(line);
    lget(c, bytes, line);
    if array {
        hex_import(c, "ecma:array", "from", 1, line);
    } else {
        c.emit_string_const("latin1", line);
        hex_import(c, "node:buffer", "toString", 2, line);
    }
    c.emit_else(line);
    c.emit_bool_const(false, line);
    if array {
        // The previous binary-string-to-array helper traps here on false.
        hex_import(c, "wasm:js-string", "length", 1, line);
        c.emit_op(Op::DROP, line);
        c.emit_i32_const(0, line);
        hex_import(c, "vybe:js-array", "newWithLength", 1, line);
    }
    c.emit_end(line);
    c.emit_end(line);
}

pub fn emit_str_rot13(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current],
            "__vybe_encoding_str_rot13",
            1,
            line,
        );
        return;
    }

    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
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
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    // charCodeAt proves an unsigned UTF-16 unit. Folding ASCII case and
    // subtracting 'a' lets one unsigned comparison recognize both A..Z/a..z.
    // No narrowing affects nonletters; their original unit reaches fromCharCode.
    lget(chunk, out_slot, line);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_i32_const(97, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, tmp_slot, line);
    chunk.emit_i32_const(25, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_if_value(line);
    lget(chunk, code_slot, line);
    chunk.emit_i32_const(13, line);
    chunk.emit_i32_const(-13, line);
    lget(chunk, tmp_slot, line);
    chunk.emit_i32_const(13, line);
    chunk.emit_op(Op::I32_LT_U, line);
    chunk.emit_op(Op::SELECT, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_else(line);
    lget(chunk, code_slot, line);
    chunk.emit_end(line);
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
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    emit_parts_join(chunk, line);
}

pub fn emit_quoted_printable_encode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current],
            "__vybe_encoding_quoted_printable_encode",
            1,
            line,
        );
        return;
    }

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
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    // need = code==61 | code>126 | (code<32 & code!=9)
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(61.0), line);
    chunk.emit_op(Op::F64_EQ, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(126.0), line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_op(Op::I32_OR, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(32.0), line);
    chunk.emit_op(Op::F64_LT, line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(9.0), line);
    chunk.emit_op(Op::F64_EQ, line);
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
    emit_append_part(chunk, out_slot, line);
    lget(chunk, hex_slot, line);
    lget(chunk, hi_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    emit_append_part(chunk, out_slot, line);
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
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    emit_parts_join(chunk, line);
}

pub fn emit_quoted_printable_decode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current],
            "__vybe_encoding_quoted_printable_decode",
            1,
            line,
        );
        return;
    }

    // length is an i32. For nonnegative lengths every cursor/advance stays
    // within it; a negative wrapped length retains the original empty result.
    // The remaining-length guard below avoids overflowing index+2.
    let c = &mut chunks[current];
    let base = c.alloc_scratch(11);
    let text = base;
    let parts = base + 1;
    let index = base + 2;
    let len = base + 3;
    let part = base + 4;
    let advance = base + 5;
    let look = base + 6;
    let first = base + 7;
    let second = base + 8;
    let hi = base + 9;
    let lo = base + 10;
    lset(c, text, line);
    emit_parts_new(c, line);
    lset(c, parts, line);
    lget(c, text, line);
    let length = c.add_import("wasm:js-string", "length");
    c.emit_call(length, 1, line);
    lset(c, len, line);
    c.emit_i32_const(0, line);
    lset(c, index, line);
    let char_at = c.add_import("ecma:string", "charAt");
    let code_at = c.add_import("wasm:js-string", "charCodeAt");
    let from_code = c.add_import("wasm:js-string", "fromCharCode");
    let done = c.emit_block(line);
    let (again, _) = c.emit_loop_s(line);
    lget(c, index, line);
    lget(c, len, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);
    lget(c, text, line);
    lget(c, index, line);
    c.emit_call(char_at, 2, line);
    lset(c, part, line);
    c.emit_i32_const(1, line);
    lset(c, advance, line);
    lget(c, text, line);
    lget(c, index, line);
    c.emit_call(code_at, 2, line);
    c.emit_i32_const(61, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_if(line);
    // Transport padding before a soft break is omitted. All code-unit reads
    // are bounded, including truncated escapes and a final '='.
    lget(c, index, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    lset(c, look, line);
    let padding_done = c.emit_block(line);
    let (padding, _) = c.emit_loop_s(line);
    lget(c, look, line);
    lget(c, len, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);
    lget(c, text, line);
    lget(c, look, line);
    c.emit_call(code_at, 2, line);
    lset(c, first, line);
    for code in [32, 9] {
        lget(c, first, line);
        c.emit_i32_const(code, line);
        c.emit_op(Op::I32_EQ, line);
    }
    c.emit_op(Op::I32_OR, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    lget(c, look, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    lset(c, look, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(padding);
    c.emit_end(line);
    c.patch_block(padding_done);
    lget(c, look, line);
    lget(c, len, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_if(line);
    push_str(c, "", line);
    lset(c, part, line);
    lget(c, len, line);
    lget(c, index, line);
    c.emit_op(Op::I32_SUB, line);
    lset(c, advance, line);
    c.emit_else(line);
    for code in [13, 10] {
        lget(c, first, line);
        c.emit_i32_const(code, line);
        c.emit_op(Op::I32_EQ, line);
    }
    c.emit_op(Op::I32_OR, line);
    c.emit_if(line);
    push_str(c, "", line);
    lset(c, part, line);
    lget(c, look, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    lset(c, look, line);
    lget(c, first, line);
    c.emit_i32_const(13, line);
    c.emit_op(Op::I32_EQ, line);
    lget(c, look, line);
    lget(c, len, line);
    c.emit_op(Op::I32_LT_S, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    lget(c, text, line);
    lget(c, look, line);
    c.emit_call(code_at, 2, line);
    c.emit_i32_const(10, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_if(line);
    lget(c, look, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    lset(c, look, line);
    c.emit_end(line);
    c.emit_end(line);
    lget(c, look, line);
    lget(c, index, line);
    c.emit_op(Op::I32_SUB, line);
    lset(c, advance, line);
    c.emit_else(line);
    lget(c, index, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    lget(c, look, line);
    c.emit_op(Op::I32_EQ, line);
    // index+2 can overflow at the final two units of an i32-max string.
    // Remaining length is nonnegative and bounded, so compare it instead.
    lget(c, len, line);
    lget(c, index, line);
    c.emit_op(Op::I32_SUB, line);
    c.emit_i32_const(2, line);
    c.emit_op(Op::I32_GT_S, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    lget(c, text, line);
    lget(c, index, line);
    c.emit_i32_const(2, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_call(code_at, 2, line);
    lset(c, second, line);
    emit_hexval(c, first, line);
    lset(c, hi, line);
    emit_hexval(c, second, line);
    lset(c, lo, line);
    // Each digit is -1 or0..15: OR is negative iff either decode failed.
    lget(c, hi, line);
    lget(c, lo, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_if(line);
    lget(c, hi, line);
    c.emit_i32_const(4, line);
    c.emit_op(Op::I32_SHL, line);
    lget(c, lo, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_call(from_code, 1, line);
    lset(c, part, line);
    c.emit_i32_const(3, line);
    lset(c, advance, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    lget(c, parts, line);
    lget(c, part, line);
    emit_append_part(c, parts, line);
    c.emit_op(Op::DROP, line);
    lget(c, index, line);
    lget(c, advance, line);
    c.emit_op(Op::I32_ADD, line);
    lset(c, index, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(again);
    c.emit_end(line);
    c.patch_block(done);
    lget(c, parts, line);
    emit_parts_join(c, line);
}

pub fn emit_convert_uuencode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current],
            "__vybe_encoding_convert_uuencode",
            1,
            line,
        );
        return;
    }

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
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
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
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_SHL, line);
    emit_uu_div_pow2(chunk, b1_slot, 16.0, line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, c_slot, line);
    lget(chunk, out_slot, line);
    emit_uu_enc(chunk, c_slot, line);
    emit_append_part(chunk, out_slot, line);
    lset(chunk, out_slot, line);
    // c2 = (b1 mod 16)*4 + floor(b2/64)
    emit_uu_mod_pow2(chunk, b1_slot, 16.0, line);
    chunk.emit_i32_const(2, line);
    chunk.emit_op(Op::I32_SHL, line);
    emit_uu_div_pow2(chunk, b2_slot, 64.0, line);
    chunk.emit_op(Op::I32_ADD, line);
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
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    push_str(chunk, "\n", line);
    emit_append_part(chunk, out_slot, line);
    emit_parts_join(chunk, line);
}

pub fn emit_convert_uudecode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current],
            "__vybe_encoding_convert_uudecode",
            1,
            line,
        );
        return;
    }

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
    emit_uu_dec(chunk, line);
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
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, done_slot, line);
    lget(chunk, count_slot, line);
    chunk.emit_op(Op::F64_LT, line);
    let _ = chunk;
    crate::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);
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
        emit_uu_dec(chunk, line);
        lset(chunk, slot, line);
    }
    // b0 = c0*4 + floor(c1/16)
    lget(chunk, c0, line);
    chunk.emit_i32_const(2, line);
    chunk.emit_op(Op::I32_SHL, line);
    emit_uu_div_pow2(chunk, c1, 16.0, line);
    chunk.emit_op(Op::I32_OR, line);
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
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_SHL, line);
    emit_uu_div_pow2(chunk, c2, 4.0, line);
    chunk.emit_op(Op::I32_OR, line);
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
    chunk.emit_i32_const(6, line);
    chunk.emit_op(Op::I32_SHL, line);
    lget(chunk, c3, line);
    chunk.emit_op(Op::I32_OR, line);
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
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
    emit_parts_join(chunk, line);
}

pub(crate) fn build_shared_helper(name: &str) -> Chunk {
    let (emitter, argc): (fn(&mut [Chunk], usize, u8, u32), u8) = match name {
        "__stdlib_encoding_base64_decode" => (emit_base64_decode, 2),
        "__stdlib_encoding_base64_decode_lenient" => (emit_base64_decode, 1),
        "__stdlib_encoding_bin2hex" => (emit_bin2hex, 1),
        "__stdlib_encoding_hex2bin" => (emit_hex2bin, 1),
        "__stdlib_encoding_hex2bytes" => (emit_hex2bytes, 1),
        "__stdlib_encoding_str_rot13" => (emit_str_rot13, 1),
        "__stdlib_encoding_quoted_printable_encode" => (emit_quoted_printable_encode, 1),
        "__stdlib_encoding_quoted_printable_decode" => (emit_quoted_printable_decode, 1),
        "__stdlib_encoding_convert_uuencode" => (emit_convert_uuencode, 1),
        "__stdlib_encoding_convert_uudecode" => (emit_convert_uudecode, 1),
        _ => unreachable!("unknown encoding helper"),
    };
    let mut c = Chunk::new(name);
    c.arity = argc;
    c.local_count = argc as u16;
    for slot in 0..argc {
        c.emit_op_u16(Op::LOCAL_GET, slot as u16, 0);
    }
    let mut chunks = vec![c];
    emitter(&mut chunks, 0, argc, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    chunks.pop().unwrap()
}
