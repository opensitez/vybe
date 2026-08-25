//! Python `quopri` and `locale`.
//!
//! `quopri` is RFC 1521 §5.1 quoted-printable, byte for byte. `locale` reads
//! the process environment (`LC_ALL` → `LC_CTYPE` → `LANG`, CPython's own
//! order) instead of inventing a locale, so an unset environment answers the
//! same `(None, None)` CPython answers.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::{base64, collections, loops, string_encoding, tuples};

use super::adapter_util::{call_import, lget, lset, stash_exact};

/// `quopri.encodestring(data)` / `decodestring(data)`.
///
/// The transform itself is the SHARED one in
/// `primitives/string_encoding.rs` — php's `quoted_printable_encode` binds
/// the same emitter. All this adds is Python's bytes⇄str boundary: CPython's
/// `quopri` takes and returns `bytes`, so the payload is widened to the
/// binary string the shared primitive works on and narrowed back afterwards
/// (the same move `base64_adapter` makes for `b64encode`).
fn to_binary_string(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    vybe_compiler::primitives::reflection::emit_typeof(chunks, current, line);
    chunks[current].emit_string_const("string", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], slot, line);
    chunks[current].emit_else(line);
    base64::emit_byte_array_slot_to_binary_string(chunks, current, Some(slot), None, None, line);
    chunks[current].emit_end(line);
}

/// A binary string on the stack → the `bytes` Python expects back.
fn to_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let text = chunks[current].alloc_scratch(2);
    let i = text + 1;
    let out = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], text, line);
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], out, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);
    let loop_id = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], text, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], text, line);
    lget(&mut chunks[current], i, line);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    loops::emit_loop_end(chunks, current, loop_id, line);
    lget(&mut chunks[current], out, line);
    call_import(chunks, current, "ecma:uint8array", "new", 1, line);
}

/// `quopri.encodestring(data)` — RFC 1521 §5.1.
pub fn emit_encodestring(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    to_binary_string(chunks, current, base, line);
    string_encoding::emit_quoted_printable_encode(chunks, current, 1, line);
    to_bytes(chunks, current, line);
}

/// `quopri.decodestring(data)` — the inverse.
pub fn emit_decodestring(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    to_binary_string(chunks, current, base, line);
    string_encoding::emit_quoted_printable_decode(chunks, current, 1, line);
    to_bytes(chunks, current, line);
}

/// `locale.getlocale()` → `(language, encoding)` parsed out of the
/// environment, exactly the categories CPython consults and in the same
/// order. An unset environment gives `(None, None)`, which is also what
/// CPython answers there.
pub fn emit_getlocale(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let raw = chunks[current].alloc_scratch(2);
    let cut = raw + 1;
    push_locale_env(chunks, current, line);
    lset(&mut chunks[current], raw, line);

    lget(&mut chunks[current], raw, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_else(line);
    // `en_US.UTF-8` → ("en_US", "UTF-8"); a bare `C` has no encoding half.
    lget(&mut chunks[current], raw, line);
    chunks[current].emit_string_const(".", line);
    call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    lset(&mut chunks[current], cut, line);
    lget(&mut chunks[current], cut, line);
    call_import(chunks, current, "wasm:js-number", "toI32", 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], raw, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], raw, line);
    chunks[current].emit_i32_const(0, line);
    lget(&mut chunks[current], cut, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    lget(&mut chunks[current], raw, line);
    lget(&mut chunks[current], cut, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lget(&mut chunks[current], raw, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    tuples::emit_tuple(chunks, current, 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `LC_ALL` → `LC_CTYPE` → `LANG`, the first that is set; null when none is.
///
/// `wasi:cli/environment.get-environment` is nullary and returns
/// `list<tuple<string, string>>` — there is no single-key lookup in the
/// interface, so a caller scans the list, which is what this does.
fn push_locale_env(chunks: &mut [Chunk], current: usize, line: u32) {
    let env = chunks[current].alloc_scratch(2);
    let found = env + 1;
    call_import(chunks, current, "wasi:cli/environment", "get-environment", 0, line);
    lset(&mut chunks[current], env, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    lset(&mut chunks[current], found, line);

    // Lowest priority first, so a later match overwrites an earlier one.
    for key in ["LANG", "LC_CTYPE", "LC_ALL"] {
        let i = chunks[current].alloc_scratch(2);
        let pair = i + 1;
        chunks[current].emit_i32_const(0, line);
        lset(&mut chunks[current], i, line);
        let loop_id = loops::emit_loop_start(chunks, current, line);
        lget(&mut chunks[current], i, line);
        lget(&mut chunks[current], env, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        loops::emit_loop_cond(chunks, current, line);

        lget(&mut chunks[current], env, line);
        lget(&mut chunks[current], i, line);
        collections::emit_get(chunks, current, line);
        lset(&mut chunks[current], pair, line);
        lget(&mut chunks[current], pair, line);
        chunks[current].emit_i32_const(0, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_string_const(key, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        lget(&mut chunks[current], pair, line);
        chunks[current].emit_i32_const(1, line);
        collections::emit_get(chunks, current, line);
        lset(&mut chunks[current], found, line);
        chunks[current].emit_end(line);

        lget(&mut chunks[current], i, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        lset(&mut chunks[current], i, line);
        loops::emit_loop_end(chunks, current, loop_id, line);
    }
    lget(&mut chunks[current], found, line);
}

/// `locale.getpreferredencoding()` — the encoding half of the same
/// environment, defaulting to the UTF-8 every WASI string already is.
pub fn emit_getpreferredencoding(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_string_const("UTF-8", line);
}
