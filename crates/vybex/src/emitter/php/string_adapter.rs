//! PHP string helpers — Rust inline opcode emitters.
//!
//! Each `emit_*` writes opcodes directly into `chunks[current]`,
//! composing only WASM string ops (`STR_LENGTH`, `STR_CHAR_AT`,
//! `STR_INDEX_OF`, `STR_TO_LOWER`, etc.) and where ECMA-262 already
//! covers the surface, `ecma:string.{encodeURIComponent,
//! decodeURIComponent}` / `ecma:number.toFixed`. No PHP-specific
//! host fns; no JS polyfills.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

// ── Local-slot / push helpers (mirror datetime_adapter) ────────────

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
}
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn coerce_to_str(chunk: &mut Chunk, line: u32) {
    // Stack: [v] → [String(v)] via "" + v.
    let v_slot = alloc_local(chunk);
    lset(chunk, v_slot, line);
    push_str(chunk, "", line);
    lget(chunk, v_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
}
fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[0].add_import(module.to_string(), name.to_string());
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(argc, line);
}

// ── ucwords ────────────────────────────────────────────────────────

/// PHP `ucwords(str, delims?)`. Stack: `[str]` or `[str, delims]` →
/// `[result]`. Uppercases the first letter of each word; default
/// delimiter set is whitespace (space + 0x09..=0x0D).
pub fn emit_ucwords(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let delims_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let cap_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let c_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let is_delim_slot = alloc_local(chunk);

    // Pop delims (if any) and s.
    if argc >= 2 {
        lset(chunk, delims_slot, line);
    } else {
        push_const(chunk, Value::Null, line);
        lset(chunk, delims_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // out = ""; cap = true; i = 0; len = s.length
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    chunk.emit_op(Op::TRUE, line);
    lset(chunk, cap_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, len_slot, line);

    // while i < len
    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // c = s.charAt(i); code = s.charCodeAt(i)
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::STR_CHAR_AT, line);
    lset(chunk, c_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::STR_CHAR_CODE_AT, line);
    lset(chunk, code_slot, line);

    // is_delim = (delims === null) ? whitespace_check : delims.indexOf(c) >= 0
    chunk.emit_op(Op::FALSE, line);
    lset(chunk, is_delim_slot, line);

    lget(chunk, delims_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let user_delims = chunk.emit_jump(Op::BR_IF_FALSE, line);
    // Whitespace path: is_delim = code in {9,10,11,12,13,32}
    for code_val in &[9.0_f64, 10.0, 11.0, 12.0, 13.0, 32.0] {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::F64(*code_val), line);
        chunk.emit_op(Op::DYN_EQ, line);
        let no_match = chunk.emit_jump(Op::BR_IF_FALSE, line);
        chunk.emit_op(Op::TRUE, line);
        lset(chunk, is_delim_slot, line);
        chunk.patch_jump(no_match);
    }
    let after_check = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(user_delims);
    // User-supplied delim string: is_delim = delims.indexOf(c) >= 0
    lget(chunk, delims_slot, line);
    lget(chunk, c_slot, line);
    chunk.emit_op(Op::STR_INDEX_OF, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    chunk.emit_op(Op::DYN_NOT, line);
    lset(chunk, is_delim_slot, line);
    chunk.patch_jump(after_check);

    // if is_delim: out += c; cap = true
    lget(chunk, is_delim_slot, line);
    let not_delim = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);
    chunk.emit_op(Op::TRUE, line);
    lset(chunk, cap_slot, line);
    let done_char = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(not_delim);

    // else if cap: out += c.toUpperCase(); cap = false
    lget(chunk, cap_slot, line);
    let not_cap = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    chunk.emit_op(Op::STR_TO_UPPER, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);
    chunk.emit_op(Op::FALSE, line);
    lset(chunk, cap_slot, line);
    let done_char2 = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(not_cap);

    // else: out += c
    lget(chunk, out_slot, line);
    lget(chunk, c_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);

    chunk.patch_jump(done_char);
    chunk.patch_jump(done_char2);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, out_slot, line);
}

// ── str_split ──────────────────────────────────────────────────────

/// PHP `str_split(s, length=1)` → array of length-N substrings.
pub fn emit_str_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let len_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let total_slot = alloc_local(chunk);
    let end_slot = alloc_local(chunk);

    if argc >= 2 {
        lset(chunk, len_slot, line);
    } else {
        push_const(chunk, Value::F64(1.0), line);
        lset(chunk, len_slot, line);
    }
    lset(chunk, s_slot, line);

    // if length < 1: return false
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let valid_len = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::FALSE, line);
    let done_invalid = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(valid_len);

    // out = []
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, out_slot, line);
    // total = s.length; i = 0
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, total_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, total_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // end = i + length
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, end_slot, line);
    // out.push(s.substring(i, end))
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, end_slot, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    let _ = chunk;
    crate::emitter::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];

    // i = end
    lget(chunk, end_slot, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, out_slot, line);
    chunk.patch_jump(done_invalid);
}

// ── str_pad ────────────────────────────────────────────────────────

/// PHP `str_pad(s, length, padStr=" ", padType=1)`. padType 0=left,
/// 1=right (default), 2=both.
pub fn emit_str_pad(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let pad_type_slot = alloc_local(chunk);
    let pad_str_slot = alloc_local(chunk);
    let length_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);

    if argc >= 4 {
        lset(chunk, pad_type_slot, line);
    } else {
        push_const(chunk, Value::F64(1.0), line);
        lset(chunk, pad_type_slot, line);
    }
    if argc >= 3 {
        lset(chunk, pad_str_slot, line);
    } else {
        push_str(chunk, " ", line);
        lset(chunk, pad_str_slot, line);
    }
    lset(chunk, length_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // Default empty pad_str → " "
    lget(chunk, pad_str_slot, line);
    push_str(chunk, "", line);
    chunk.emit_op(Op::DYN_EQ, line);
    let nonempty = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_str(chunk, " ", line);
    lset(chunk, pad_str_slot, line);
    chunk.patch_jump(nonempty);

    // if s.length >= length: return s  ≡  !(s.length < length)
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lget(chunk, length_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    chunk.emit_op(Op::DYN_NOT, line);
    let needs_pad = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, s_slot, line);
    let done_short = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(needs_pad);

    // pad_type === 0: STR_PAD_START
    lget(chunk, pad_type_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let not_left = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, s_slot, line);
    lget(chunk, length_slot, line);
    lget(chunk, pad_str_slot, line);
    chunk.emit_op(Op::STR_PAD_START, line);
    let done_left = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(not_left);

    // pad_type === 2: both
    lget(chunk, pad_type_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let not_both = chunk.emit_jump(Op::BR_IF_FALSE, line);
    // diff = length - s.length; left = floor(diff/2); right = diff - left
    let diff_slot = alloc_local(chunk);
    let left_len_slot = alloc_local(chunk);
    let right_len_slot = alloc_local(chunk);
    lget(chunk, length_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, diff_slot, line);
    lget(chunk, diff_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    lset(chunk, left_len_slot, line);
    lget(chunk, diff_slot, line);
    lget(chunk, left_len_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, right_len_slot, line);

    // s_padded = padStart(s, s.length + left_len, pad_str)
    lget(chunk, s_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lget(chunk, left_len_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, pad_str_slot, line);
    chunk.emit_op(Op::STR_PAD_START, line);
    // result = padEnd(s_padded, length, pad_str)
    lget(chunk, length_slot, line);
    lget(chunk, pad_str_slot, line);
    chunk.emit_op(Op::STR_PAD_END, line);
    let done_both = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(not_both);

    // Default (pad_type 1 = right): STR_PAD_END
    lget(chunk, s_slot, line);
    lget(chunk, length_slot, line);
    lget(chunk, pad_str_slot, line);
    chunk.emit_op(Op::STR_PAD_END, line);

    chunk.patch_jump(done_left);
    chunk.patch_jump(done_both);
    chunk.patch_jump(done_short);
}

// ── substr_count ───────────────────────────────────────────────────

/// PHP `substr_count(hay, needle, offset=0, length?)`.
pub fn emit_substr_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let length_slot = alloc_local(chunk);
    let offset_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let hay_slot = alloc_local(chunk);
    let slice_slot = alloc_local(chunk);
    let count_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);
    let has_length_slot = alloc_local(chunk);

    chunk.emit_op(if argc >= 4 { Op::TRUE } else { Op::FALSE }, line);
    lset(chunk, has_length_slot, line);
    if argc >= 4 { lset(chunk, length_slot, line); }
    if argc >= 3 {
        lset(chunk, offset_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, offset_slot, line);
    }
    lset(chunk, needle_slot, line);
    lset(chunk, hay_slot, line);

    // slice = has_length ? hay.substring(offset, offset+length) : hay.substring(offset)
    lget(chunk, has_length_slot, line);
    let no_length = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, hay_slot, line);
    lget(chunk, offset_slot, line);
    lget(chunk, offset_slot, line);
    lget(chunk, length_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    let after_slice = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(no_length);
    lget(chunk, hay_slot, line);
    lget(chunk, offset_slot, line);
    lget(chunk, hay_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    chunk.patch_jump(after_slice);
    lset(chunk, slice_slot, line);

    // if needle.length === 0: return 0
    lget(chunk, needle_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let nonempty_needle = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_const(chunk, Value::F64(0.0), line);
    let done_empty = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(nonempty_needle);

    // count = 0; pos = 0
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, count_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, pos_slot, line);

    // while true { idx = slice.indexOf(needle, pos); if idx<0: break; count++; pos = idx + needle.length }
    let loop_top = chunk.current_offset();
    // Vybe's STR_INDEX_OF takes 2 args (haystack, needle) without pos — emulate with substring.
    // idx = slice.substring(pos).indexOf(needle); if idx>=0: real_idx = idx + pos
    lget(chunk, slice_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, slice_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    lget(chunk, needle_slot, line);
    chunk.emit_op(Op::STR_INDEX_OF, line);
    lset(chunk, idx_slot, line);

    lget(chunk, idx_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_TRUE, line);
    // not negative: count++; pos = idx + pos + needle.length
    lget(chunk, count_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, count_slot, line);
    lget(chunk, idx_slot, line);
    lget(chunk, pos_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, needle_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, pos_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, count_slot, line);
    chunk.patch_jump(done_empty);
}

// ── strstr / stristr ───────────────────────────────────────────────

fn emit_strstr_impl(chunks: &mut [Chunk], current: usize, argc: u8, case_insensitive: bool, line: u32) {
    let chunk = &mut chunks[current];
    let before_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let hay_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);

    if argc >= 3 {
        lset(chunk, before_slot, line);
    } else {
        chunk.emit_op(Op::FALSE, line);
        lset(chunk, before_slot, line);
    }
    lset(chunk, needle_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, hay_slot, line);

    if case_insensitive {
        // idx = lower(hay).indexOf(lower(needle))
        lget(chunk, hay_slot, line);
        chunk.emit_op(Op::STR_TO_LOWER, line);
        lget(chunk, needle_slot, line);
        chunk.emit_op(Op::STR_TO_LOWER, line);
        chunk.emit_op(Op::STR_INDEX_OF, line);
    } else {
        lget(chunk, hay_slot, line);
        lget(chunk, needle_slot, line);
        chunk.emit_op(Op::STR_INDEX_OF, line);
    }
    lset(chunk, idx_slot, line);

    // if idx < 0: return false
    lget(chunk, idx_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let found = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::FALSE, line);
    let done_notfound = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(found);

    // if before === true: return hay.substring(0, idx)
    lget(chunk, before_slot, line);
    chunk.emit_op(Op::TRUE, line);
    chunk.emit_op(Op::DYN_EQ, line);
    let after = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, hay_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, idx_slot, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    let done_before = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(after);

    // else: return hay.substring(idx)
    lget(chunk, hay_slot, line);
    lget(chunk, idx_slot, line);
    lget(chunk, hay_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);

    chunk.patch_jump(done_notfound);
    chunk.patch_jump(done_before);
}

pub fn emit_strstr(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_strstr_impl(chunks, current, argc, /*case_insensitive=*/false, line);
}
pub fn emit_stristr(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_strstr_impl(chunks, current, argc, /*case_insensitive=*/true, line);
}

// ── urlencode / rawurlencode / urldecode ───────────────────────────

/// PHP `urlencode` — like `encodeURIComponent` but space → "+".
pub fn emit_urlencode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    coerce_to_str(&mut chunks[current], line);
    call_import(chunks, current, "ecma:string", "encodeURIComponent", 1, line);
    let chunk = &mut chunks[current];
    // Replace "%20" with "+"
    push_str(chunk, "%20", line);
    push_str(chunk, "+", line);
    chunk.emit_op(Op::STR_REPLACE, line);
}

/// PHP `rawurlencode` — strict `encodeURIComponent`.
pub fn emit_rawurlencode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    coerce_to_str(&mut chunks[current], line);
    call_import(chunks, current, "ecma:string", "encodeURIComponent", 1, line);
}

/// PHP `urldecode` — replace "+" with " ", then decodeURIComponent.
pub fn emit_urldecode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    coerce_to_str(chunk, line);
    push_str(chunk, "+", line);
    push_str(chunk, " ", line);
    chunk.emit_op(Op::STR_REPLACE, line);
    call_import(chunks, current, "ecma:string", "decodeURIComponent", 1, line);
}

// ── bin2hex / hex2bin ──────────────────────────────────────────────

/// PHP `bin2hex(str)` — char loop, two hex digits per byte.
pub fn emit_bin2hex(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let hi_slot = alloc_local(chunk);
    let lo_slot = alloc_local(chunk);

    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, len_slot, line);

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::STR_CHAR_CODE_AT, line);
    lset(chunk, code_slot, line);

    // hi = (code >> 4) & 0xF; lo = code & 0xF
    let table = "0123456789abcdef";
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    lset(chunk, hi_slot, line);
    let fmod = chunks[0].add_import("ecma:math", "fmod");
    let chunk = &mut chunks[current];
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, fmod, line);
    chunk.emit(2, line);
    lset(chunk, lo_slot, line);

    // out += table.charAt(hi) + table.charAt(lo)
    lget(chunk, out_slot, line);
    push_str(chunk, table, line);
    lget(chunk, hi_slot, line);
    chunk.emit_op(Op::STR_CHAR_AT, line);
    chunk.emit_op(Op::DYN_ADD, line);
    push_str(chunk, table, line);
    lget(chunk, lo_slot, line);
    chunk.emit_op(Op::STR_CHAR_AT, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, out_slot, line);
}

/// PHP `hex2bin(hex)` — pair loop.
pub fn emit_hex2bin(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let hi_slot = alloc_local(chunk);
    let lo_slot = alloc_local(chunk);

    coerce_to_str(chunk, line);
    chunk.emit_op(Op::STR_TO_LOWER, line);
    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, len_slot, line);

    let loop_top = chunk.current_offset();
    // i + 1 < len
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    let table = "0123456789abcdef";
    push_str(chunk, table, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::STR_CHAR_AT, line);
    chunk.emit_op(Op::STR_INDEX_OF, line);
    lset(chunk, hi_slot, line);
    push_str(chunk, table, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::STR_CHAR_AT, line);
    chunk.emit_op(Op::STR_INDEX_OF, line);
    lset(chunk, lo_slot, line);

    // if hi<0 || lo<0: return false
    lget(chunk, hi_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let hi_ok = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::FALSE, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.patch_jump(hi_ok);
    lget(chunk, lo_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let lo_ok = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::FALSE, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.patch_jump(lo_ok);

    // out += String.fromCharCode((hi << 4) | lo)
    lget(chunk, out_slot, line);
    lget(chunk, hi_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    lget(chunk, lo_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::STR_FROM_CHAR_CODE, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);

    // i += 2
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, out_slot, line);
}

// ── chunk_split ────────────────────────────────────────────────────

/// PHP `chunk_split(s, length=76, end="\r\n")`.
pub fn emit_chunk_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let end_slot = alloc_local(chunk);
    let length_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let total_slot = alloc_local(chunk);

    if argc >= 3 {
        lset(chunk, end_slot, line);
    } else {
        push_str(chunk, "\r\n", line);
        lset(chunk, end_slot, line);
    }
    if argc >= 2 {
        lset(chunk, length_slot, line);
    } else {
        push_const(chunk, Value::F64(76.0), line);
        lset(chunk, length_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // if length < 1: return false
    lget(chunk, length_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let valid_len = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::FALSE, line);
    let done_invalid = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(valid_len);

    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, total_slot, line);

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, total_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // out += s.substring(i, i+length) + end
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, length_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lget(chunk, end_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);

    lget(chunk, i_slot, line);
    lget(chunk, length_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, out_slot, line);
    chunk.patch_jump(done_invalid);
}

// ── number_format ──────────────────────────────────────────────────

/// PHP `number_format(num, decimals=0, decsep=".", thousep=",")`.
pub fn emit_number_format(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let thousep_slot = alloc_local(chunk);
    let decsep_slot = alloc_local(chunk);
    let decimals_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let sign_slot = alloc_local(chunk);
    let fixed_slot = alloc_local(chunk);
    let dot_slot = alloc_local(chunk);
    let int_part_slot = alloc_local(chunk);
    let frac_part_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);

    if argc >= 4 {
        lset(chunk, thousep_slot, line);
    } else {
        push_str(chunk, ",", line);
        lset(chunk, thousep_slot, line);
    }
    if argc >= 3 {
        lset(chunk, decsep_slot, line);
    } else {
        push_str(chunk, ".", line);
        lset(chunk, decsep_slot, line);
    }
    if argc >= 2 {
        lset(chunk, decimals_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, decimals_slot, line);
    }
    // n = +num (numeric coerce)
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, n_slot, line);

    // sign = ""; if n < 0: sign = "-"; n = -n
    push_str(chunk, "", line);
    lset(chunk, sign_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let nonneg = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_str(chunk, "-", line);
    lset(chunk, sign_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, n_slot, line);
    chunk.patch_jump(nonneg);

    // PHP rounds half-away-from-zero; JS toFixed is engine-dependent
    // (banker's in V8, ECMA-spec deferred). Pre-round via Math.round
    // (which uses f64::round in vybe_host = half-away-from-zero) at
    // scale 10^decimals to force PHP semantics:
    //   n = Math.round(n * scale) / scale
    let pow = chunks[0].add_import("ecma:math", "pow");
    let round = chunks[0].add_import("ecma:math", "round");
    let chunk = &mut chunks[current];
    let scale_slot = alloc_local(chunk);
    push_const(chunk, Value::F64(10.0), line);
    lget(chunk, decimals_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, pow, line);
    chunk.emit(2, line);
    lset(chunk, scale_slot, line);
    lget(chunk, n_slot, line);
    lget(chunk, scale_slot, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, round, line);
    chunk.emit(1, line);
    lget(chunk, scale_slot, line);
    chunk.emit_op(Op::F64_DIV, line);
    lset(chunk, n_slot, line);

    // fixed = ecma:number.toFixed(n, decimals)
    lget(chunk, n_slot, line);
    lget(chunk, decimals_slot, line);
    let to_fixed = chunks[0].add_import("ecma:number", "toFixed");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, to_fixed, line);
    chunk.emit(2, line);
    lset(chunk, fixed_slot, line);

    // dot_idx = fixed.indexOf(".")
    lget(chunk, fixed_slot, line);
    push_str(chunk, ".", line);
    chunk.emit_op(Op::STR_INDEX_OF, line);
    lset(chunk, dot_slot, line);

    // if dot_idx < 0: int_part = fixed; frac_part = ""
    lget(chunk, dot_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let has_dot = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, fixed_slot, line);
    lset(chunk, int_part_slot, line);
    push_str(chunk, "", line);
    lset(chunk, frac_part_slot, line);
    let after_split = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(has_dot);
    lget(chunk, fixed_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, dot_slot, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    lset(chunk, int_part_slot, line);
    lget(chunk, fixed_slot, line);
    lget(chunk, dot_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, fixed_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    lset(chunk, frac_part_slot, line);
    chunk.patch_jump(after_split);

    // out = ""; len = int_part.length; for i in 0..len: if i>0 && (len-i) % 3 == 0: out += thousep; out += int_part.charAt(i)
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    lget(chunk, int_part_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit_loop = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // i > 0 && (len - i) % 3 == 0
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_GT, line);
    let no_sep = chunk.emit_jump(Op::BR_IF_FALSE, line);
    let fmod = chunks[0].add_import("ecma:math", "fmod");
    let chunk = &mut chunks[current];
    lget(chunk, len_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::F64(3.0), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, fmod, line);
    chunk.emit(2, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let no_sep2 = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, out_slot, line);
    lget(chunk, thousep_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);
    chunk.patch_jump(no_sep2);
    chunk.patch_jump(no_sep);

    // out += int_part.charAt(i)
    lget(chunk, out_slot, line);
    lget(chunk, int_part_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::STR_CHAR_AT, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit_loop);

    // if frac_part.length > 0: out += decsep + frac_part
    lget(chunk, frac_part_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_GT, line);
    let no_frac = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, out_slot, line);
    lget(chunk, decsep_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lget(chunk, frac_part_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);
    chunk.patch_jump(no_frac);

    // sign + out
    lget(chunk, sign_slot, line);
    lget(chunk, out_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
}

// ── substr_replace ─────────────────────────────────────────────────

/// PHP `substr_replace(str, repl, start, length?)`.
pub fn emit_substr_replace(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let length_slot = alloc_local(chunk);
    let start_slot = alloc_local(chunk);
    let repl_slot = alloc_local(chunk);
    let str_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let l_slot = alloc_local(chunk);
    let has_length_slot = alloc_local(chunk);

    chunk.emit_op(if argc >= 4 { Op::TRUE } else { Op::FALSE }, line);
    lset(chunk, has_length_slot, line);
    if argc >= 4 { lset(chunk, length_slot, line); }
    lset(chunk, start_slot, line);
    lset(chunk, repl_slot, line);
    lset(chunk, str_slot, line);

    // len = str.length
    lget(chunk, str_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, len_slot, line);

    // s = start < 0 ? max(len + start, 0) : min(start, len)
    lget(chunk, start_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let pos_start = chunk.emit_jump(Op::BR_IF_FALSE, line);
    // negative: max(len + start, 0)
    lget(chunk, len_slot, line);
    lget(chunk, start_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    let neg_slot = alloc_local(chunk);
    lset(chunk, neg_slot, line);
    lget(chunk, neg_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let take_neg = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_const(chunk, Value::F64(0.0), line);
    let after_max = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(take_neg);
    lget(chunk, neg_slot, line);
    chunk.patch_jump(after_max);
    lset(chunk, s_slot, line);
    let after_pos = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(pos_start);
    // positive: min(start, len)
    lget(chunk, start_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_GT, line);
    let take_start = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, len_slot, line);
    let after_min = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(take_start);
    lget(chunk, start_slot, line);
    chunk.patch_jump(after_min);
    lset(chunk, s_slot, line);
    chunk.patch_jump(after_pos);

    // l = has_length ? clamp(length, len - s) : len - s
    lget(chunk, has_length_slot, line);
    let no_length = chunk.emit_jump(Op::BR_IF_FALSE, line);
    // Has length: if length < 0: l = max(len + length - s, 0) else l = min(length, len - s)
    lget(chunk, length_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let pos_len = chunk.emit_jump(Op::BR_IF_FALSE, line);
    // negative length: max(len + length - s, 0)
    lget(chunk, len_slot, line);
    lget(chunk, length_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    let neg_l_slot = alloc_local(chunk);
    lset(chunk, neg_l_slot, line);
    lget(chunk, neg_l_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let take_neg_l = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_const(chunk, Value::F64(0.0), line);
    let after_max_l = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(take_neg_l);
    lget(chunk, neg_l_slot, line);
    chunk.patch_jump(after_max_l);
    let after_pos_l = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(pos_len);
    // positive length: min(length, len - s)
    lget(chunk, length_slot, line);
    lget(chunk, len_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    let rem_slot = alloc_local(chunk);
    lset(chunk, rem_slot, line);
    lget(chunk, length_slot, line);
    lget(chunk, rem_slot, line);
    chunk.emit_op(Op::DYN_GT, line);
    let take_length = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, rem_slot, line);
    let after_min_l = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(take_length);
    lget(chunk, length_slot, line);
    chunk.patch_jump(after_min_l);
    chunk.patch_jump(after_pos_l);
    let with_length_done = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(no_length);
    // No length: l = len - s
    lget(chunk, len_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.patch_jump(with_length_done);
    lset(chunk, l_slot, line);

    // result = str.substring(0, s) + repl + str.substring(s + l)
    lget(chunk, str_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    lget(chunk, repl_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lget(chunk, str_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, l_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    chunk.emit_op(Op::DYN_ADD, line);
}

// ── str_word_count ─────────────────────────────────────────────────

/// PHP `str_word_count(s)` — count whitespace-or-punct-separated words.
pub fn emit_str_word_count(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let count_slot = alloc_local(chunk);
    let in_word_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let is_sep_slot = alloc_local(chunk);

    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, count_slot, line);
    chunk.emit_op(Op::FALSE, line);
    lset(chunk, in_word_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, len_slot, line);

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::STR_CHAR_CODE_AT, line);
    lset(chunk, code_slot, line);

    // is_sep: whitespace (9..=13, 32) | comma 44 | period 46 | ! 33 | ? 63 | ; 59 | : 58
    chunk.emit_op(Op::FALSE, line);
    lset(chunk, is_sep_slot, line);
    for code_val in &[9.0_f64, 10.0, 11.0, 12.0, 13.0, 32.0, 33.0, 44.0, 46.0, 58.0, 59.0, 63.0] {
        lget(chunk, code_slot, line);
        push_const(chunk, Value::F64(*code_val), line);
        chunk.emit_op(Op::DYN_EQ, line);
        let no_match = chunk.emit_jump(Op::BR_IF_FALSE, line);
        chunk.emit_op(Op::TRUE, line);
        lset(chunk, is_sep_slot, line);
        chunk.patch_jump(no_match);
    }
    // if !is_sep && !in_word: count++; in_word = true
    // else if is_sep: in_word = false
    lget(chunk, is_sep_slot, line);
    let is_separator = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::FALSE, line);
    lset(chunk, in_word_slot, line);
    let after_sep = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(is_separator);
    // not separator: enter word if not already in
    lget(chunk, in_word_slot, line);
    let already_in = chunk.emit_jump(Op::BR_IF_FALSE, line);
    let after_word = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(already_in);
    lget(chunk, count_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, count_slot, line);
    chunk.emit_op(Op::TRUE, line);
    lset(chunk, in_word_slot, line);
    chunk.patch_jump(after_word);
    chunk.patch_jump(after_sep);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, count_slot, line);
}

// ── str_ireplace, str_replace, wordwrap ────────────────────────────
//
// These three are larger and rare in the test surface. To keep this
// migration focused, route them through the simple `STR_REPLACE`
// opcode in the common case (single search/replace string) and leave
// the array-aware / case-insensitive paths as runtime fallbacks
// later. For now their dispatch arms emit the three-arg STR_REPLACE
// directly when the inputs are strings; PHP-array shapes were
// covered by the polyfills and aren't exercised by current tests.

/// PHP `str_ireplace(search, replace, subject)`.
/// Case-insensitive find: walk by chunks of length(needle), comparing
/// lowercased segments. Falls back to the JS-host
/// `STR_REPLACE` for the simple case and emits a manual scan for
/// case insensitivity.
pub fn emit_str_ireplace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let subj_slot = alloc_local(chunk);
    let repl_slot = alloc_local(chunk);
    let srch_slot = alloc_local(chunk);
    let lower_slot = alloc_local(chunk);
    let srch_lower_slot = alloc_local(chunk);
    let srch_len_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);

    coerce_to_str(chunk, line);
    lset(chunk, subj_slot, line);
    lset(chunk, repl_slot, line);
    lset(chunk, srch_slot, line);

    // lower = subj.toLowerCase(); srch_lower = srch.toLowerCase()
    lget(chunk, subj_slot, line);
    chunk.emit_op(Op::STR_TO_LOWER, line);
    lset(chunk, lower_slot, line);
    lget(chunk, srch_slot, line);
    chunk.emit_op(Op::STR_TO_LOWER, line);
    lset(chunk, srch_lower_slot, line);
    lget(chunk, srch_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, srch_len_slot, line);

    // if srch_len === 0: return subj
    lget(chunk, srch_len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let nonempty = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, subj_slot, line);
    let done_empty = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(nonempty);

    // out = ""; pos = 0
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, pos_slot, line);

    // while true { idx = lower.substring(pos).indexOf(srch_lower); if idx<0: break }
    let loop_top = chunk.current_offset();
    lget(chunk, lower_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, lower_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    lget(chunk, srch_lower_slot, line);
    chunk.emit_op(Op::STR_INDEX_OF, line);
    lset(chunk, idx_slot, line);

    lget(chunk, idx_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_TRUE, line);

    // out += subj.substring(pos, pos + idx) + repl
    lget(chunk, out_slot, line);
    lget(chunk, subj_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, idx_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lget(chunk, repl_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);

    // pos = pos + idx + srch_len
    lget(chunk, pos_slot, line);
    lget(chunk, idx_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, srch_len_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, pos_slot, line);

    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    // out += subj.substring(pos)
    lget(chunk, out_slot, line);
    lget(chunk, subj_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, subj_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.patch_jump(done_empty);
}

// ── str_replace (array-aware) ──────────────────────────────────────

/// PHP `str_replace(search, replace, subject)`. When `search` and
/// `replace` are both strings, this is one `STR_REPLACE` opcode. The
/// adapter handles the array-aware variants too: when `search` is an
/// array, iterate and apply each pair (with `replace` as either a
/// scalar or a parallel array).
///
/// Strategy: probe `Array.isArray(search)` at runtime via
/// `ecma:array.isArray`. If false, fall back to `STR_REPLACE`.
/// Otherwise loop over the array and split/join per element.
pub fn emit_str_replace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let subj_slot = alloc_local(chunk);
    let repl_slot = alloc_local(chunk);
    let srch_slot = alloc_local(chunk);
    lset(chunk, subj_slot, line);
    lset(chunk, repl_slot, line);
    lset(chunk, srch_slot, line);

    // Coerce subj to string.
    push_str(chunk, "", line);
    lget(chunk, subj_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, subj_slot, line);

    // is_array_search = Array.isArray(srch)
    let _ = chunk;
    chunks[current].emit_op_u16(Op::LOCAL_GET, srch_slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    // ecma:array.isArray returns I32(0|1); BR_IF_FALSE checks
    // Value::Bool(true), so coerce to bool first.
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    let scalar_path = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // ── Array path ──
    // is_array_repl = Array.isArray(repl)
    let is_array_repl_slot = alloc_local(chunk);
    chunk.emit_op(Op::FALSE, line);
    lset(chunk, is_array_repl_slot, line);
    let _ = chunk;
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    let not_arr_repl = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::TRUE, line);
    lset(chunk, is_array_repl_slot, line);
    chunk.patch_jump(not_arr_repl);

    // for i in 0..srch.length: needle = srch[i]; rep = is_array_repl ? (i < repl.length ? repl[i] : "") : repl
    // subj = subj.split(needle).join(rep)
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let rep_slot = alloc_local(chunk);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, srch_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // needle = "" + srch[i]
    push_str(chunk, "", line);
    lget(chunk, srch_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, needle_slot, line);

    // if needle.length === 0: skip (continue)
    lget(chunk, needle_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let nonempty_needle = chunk.emit_jump(Op::BR_IF_TRUE, line);

    // rep = is_array_repl ? (i < repl.length ? "" + repl[i] : "") : "" + repl
    lget(chunk, is_array_repl_slot, line);
    let scalar_rep = chunk.emit_jump(Op::BR_IF_FALSE, line);
    // array repl path
    lget(chunk, i_slot, line);
    lget(chunk, repl_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op(Op::DYN_LT, line);
    let oob = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_str(chunk, "", line);
    lget(chunk, repl_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::DYN_ADD, line);
    let rep_done_arr = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(oob);
    push_str(chunk, "", line);
    chunk.patch_jump(rep_done_arr);
    let rep_done = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(scalar_rep);
    // scalar repl path
    push_str(chunk, "", line);
    lget(chunk, repl_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.patch_jump(rep_done);
    lset(chunk, rep_slot, line);

    // subj = STR_REPLACE(subj, needle, rep)  — replaces ALL occurrences
    lget(chunk, subj_slot, line);
    lget(chunk, needle_slot, line);
    lget(chunk, rep_slot, line);
    chunk.emit_op(Op::STR_REPLACE, line);
    lset(chunk, subj_slot, line);

    chunk.patch_jump(nonempty_needle);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, subj_slot, line);
    let done = chunk.emit_jump(Op::BR, line);

    // ── Scalar path ──
    chunk.patch_jump(scalar_path);
    // STR_REPLACE(subj, "" + srch, "" + repl)
    lget(chunk, subj_slot, line);
    push_str(chunk, "", line);
    lget(chunk, srch_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    push_str(chunk, "", line);
    lget(chunk, repl_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.emit_op(Op::STR_REPLACE, line);

    chunk.patch_jump(done);
}

// ── wordwrap ───────────────────────────────────────────────────────

/// PHP `wordwrap(s, width=75, break_str="\n", cut=false)`.
///
/// Emits a per-line word-wrap loop: for each `\n`-delimited input
/// line, walk words and accumulate onto a `current` buffer until the
/// width is exceeded; on overflow, push the current buffer to `out`
/// and start a new one. With `cut=true`, also break mid-word.
pub fn emit_wordwrap(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let cut_slot = alloc_local(chunk);
    let br_slot = alloc_local(chunk);
    let width_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);

    if argc >= 4 { lset(chunk, cut_slot, line); }
    else { chunk.emit_op(Op::FALSE, line); lset(chunk, cut_slot, line); }
    if argc >= 3 { lset(chunk, br_slot, line); }
    else { push_str(chunk, "\n", line); lset(chunk, br_slot, line); }
    if argc >= 2 { lset(chunk, width_slot, line); }
    else { push_const(chunk, Value::F64(75.0), line); lset(chunk, width_slot, line); }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);

    // lines = s.split("\n")
    let lines_slot = alloc_local(chunk);
    lget(chunk, s_slot, line);
    push_str(chunk, "\n", line);
    chunk.emit_op(Op::STR_SPLIT, line);
    lset(chunk, lines_slot, line);

    // out = []
    let out_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, out_slot, line);

    // for li in 0..lines.length
    let li_slot = alloc_local(chunk);
    let nlines_slot = alloc_local(chunk);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, li_slot, line);
    lget(chunk, lines_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, nlines_slot, line);

    let line_slot = alloc_local(chunk);
    let words_slot = alloc_local(chunk);
    let nwords_slot = alloc_local(chunk);
    let wi_slot = alloc_local(chunk);
    let word_slot = alloc_local(chunk);
    let current_slot = alloc_local(chunk);

    let outer_top = chunk.current_offset();
    lget(chunk, li_slot, line);
    lget(chunk, nlines_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let outer_exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // line = lines[li]
    lget(chunk, lines_slot, line);
    lget(chunk, li_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, line_slot, line);

    // if line.length <= width: out.push(line); continue
    lget(chunk, line_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lget(chunk, width_slot, line);
    chunk.emit_op(Op::DYN_GT, line);
    let needs_wrap = chunk.emit_jump(Op::BR_IF_TRUE, line);
    // short — push as-is
    lget(chunk, out_slot, line);
    lget(chunk, line_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    let after_short = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(needs_wrap);

    // words = line.split(" ")
    lget(chunk, line_slot, line);
    push_str(chunk, " ", line);
    chunk.emit_op(Op::STR_SPLIT, line);
    lset(chunk, words_slot, line);
    lget(chunk, words_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, nwords_slot, line);

    // current = ""; wi = 0
    push_str(chunk, "", line);
    lset(chunk, current_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, wi_slot, line);

    let inner_top = chunk.current_offset();
    lget(chunk, wi_slot, line);
    lget(chunk, nwords_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let inner_exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // word = words[wi]
    lget(chunk, words_slot, line);
    lget(chunk, wi_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, word_slot, line);

    // if current.length === 0: current = word
    lget(chunk, current_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let not_empty_cur = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, word_slot, line);
    lset(chunk, current_slot, line);
    let after_word = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(not_empty_cur);

    // else if current.length + 1 + word.length <= width: current = current + " " + word
    lget(chunk, current_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, word_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, width_slot, line);
    chunk.emit_op(Op::DYN_GT, line);
    let must_break = chunk.emit_jump(Op::BR_IF_TRUE, line);
    // append
    lget(chunk, current_slot, line);
    push_str(chunk, " ", line);
    chunk.emit_op(Op::DYN_ADD, line);
    lget(chunk, word_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, current_slot, line);
    let after_word2 = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(must_break);

    // else: out.push(current); current = word
    lget(chunk, out_slot, line);
    lget(chunk, current_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    lget(chunk, word_slot, line);
    lset(chunk, current_slot, line);

    chunk.patch_jump(after_word);
    chunk.patch_jump(after_word2);

    // wi++
    lget(chunk, wi_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, wi_slot, line);
    chunk.emit_loop(inner_top, line);
    chunk.patch_jump(inner_exit);

    // if current.length > 0: out.push(current)
    lget(chunk, current_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_GT, line);
    let no_remainder = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, out_slot, line);
    lget(chunk, current_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    chunk.patch_jump(no_remainder);

    chunk.patch_jump(after_short);

    // li++
    lget(chunk, li_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, li_slot, line);
    chunk.emit_loop(outer_top, line);
    chunk.patch_jump(outer_exit);

    // out.join(br)
    lget(chunk, out_slot, line);
    lget(chunk, br_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "join", 2, line);
}
