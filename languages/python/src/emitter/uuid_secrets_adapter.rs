//! Python `uuid` and `secrets` — both are randomness surfaces, both sit on
//! `web:crypto`.
//!
//! `web:crypto.getRandomValues` fills a plain `ObjectKind::Array` (the host
//! leaves a TypedArray untouched), so the byte generators build an array,
//! fill it, and convert the result to the `Uint8Array` Python calls `bytes`.

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::{base64, collections};

use super::adapter_util::{call_import, lget, lset, new_object, stash_exact, struct_set};

/// A length-`n` array of random bytes on the stack, `n` taken from `slot`.
fn push_random_array(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    collections::emit_new_with_length(chunks, current, line);
    call_import(chunks, current, "web:crypto", "getRandomValues", 1, line);
}

/// `n` bytes, defaulting to `default_len` when the call omits the count.
fn size_slot(chunks: &mut [Chunk], current: usize, argc: u8, default_len: i32, line: u32) -> u16 {
    let base = stash_exact(chunks, current, argc, 1, line);
    if argc == 0 {
        chunks[current].emit_i32_const(default_len, line);
        lset(&mut chunks[current], base, line);
    }
    base
}

/// `secrets.token_bytes(n=32)`.
pub fn emit_token_bytes(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let n = size_slot(chunks, current, argc, 32, line);
    push_random_array(chunks, current, n, line);
    call_import(chunks, current, "ecma:uint8array", "new", 1, line);
}

/// `secrets.token_hex(n=32)` → `2n` hex characters.
pub fn emit_token_hex(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let n = size_slot(chunks, current, argc, 32, line);
    push_random_array(chunks, current, n, line);
    call_import(chunks, current, "ecma:uint8array", "new", 1, line);
    call_import(chunks, current, "ecma:uint8array", "toHex", 1, line);
}

/// `secrets.token_urlsafe(n=32)` — base64 with the URL alphabet and no
/// padding, exactly what CPython's `token_urlsafe` returns.
pub fn emit_token_urlsafe(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let n = size_slot(chunks, current, argc, 32, line);
    let bytes = chunks[current].alloc_scratch(1);
    push_random_array(chunks, current, n, line);
    lset(&mut chunks[current], bytes, line);
    base64::emit_byte_array_slot_to_binary_string(chunks, current, Some(bytes), None, None, line);
    base64::emit_encode_binary_string(chunks, current, line);
    for (from, to) in [("+", "-"), ("/", "_"), ("=", "")] {
        chunks[current].emit_string_const(from, line);
        chunks[current].emit_string_const(to, line);
        call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    }
}

/// `secrets.randbelow(n)` → a uniform int in `[0, n)`.
pub fn emit_randbelow(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    call_import(chunks, current, "ecma:math", "random", 0, line);
    lget(&mut chunks[current], base, line);
    let to_f64 = chunks[current].add_import("wasm:js-number", "toF64");
    chunks[current].emit_call(to_f64, 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
}

/// `secrets.choice(seq)` → one element, chosen uniformly.
pub fn emit_choice(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let index = chunks[current].alloc_scratch(1);
    call_import(chunks, current, "ecma:math", "random", 0, line);
    lget(&mut chunks[current], base, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    lset(&mut chunks[current], index, line);
    lget(&mut chunks[current], base, line);
    lget(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
}

/// The canonical `8-4-4-4-12` text of a UUID, as `str()` renders it.
const CANONICAL_KEY: &str = "__uuid";

/// `(this) -> this.__uuid` — the ToString/Repr slot body.
fn build_str_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut helper = Chunk::new("__py_uuid_str");
    helper.arity = 1;
    helper.local_count = helper.local_count.max(1);
    helper.emit_op_u16(Op::LOCAL_GET, 0, line);
    let k = helper.add_constant(Value::String(std::sync::Arc::from(CANONICAL_KEY)));
    helper.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
    helper.emit_op(Op::RETURN, line);
    chunks.push(helper);
    chunks.len() - 1
}

/// Build the `UUID` object around a canonical string already on the stack.
fn wrap_uuid(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let str_idx = build_str_helper(chunks, line);
    let text = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], text, line);

    let chunk = &mut chunks[current];
    new_object(chunk, line);
    chunk.emit_dup(line);
    chunk.emit_string_const("UUID", line);
    struct_set(chunk, "__type", line);
    chunk.emit_dup(line);
    lget(chunk, text, line);
    struct_set(chunk, CANONICAL_KEY, line);
    // `.hex` is the same digits without the dashes.
    chunk.emit_dup(line);
    lget(chunk, text, line);
    chunk.emit_string_const("-", line);
    chunk.emit_string_const("", line);
    let replace_all = chunk.add_import("ecma:string", "replaceAll");
    chunk.emit_call(replace_all, 3, line);
    struct_set(chunk, "hex", line);
    chunk.emit_dup(line);
    lget(chunk, text, line);
    chunk.emit_string_const("urn:uuid:", line);
    let concat = chunk.add_import("wasm:js-string", "concat");
    // `concat(a, b)` — the prefix has to be the LEFT operand.
    chunk.emit_call(concat, 2, line);
    struct_set(chunk, "urn", line);

    for slot in [
        vybe_ast::ProtocolSlot::ToString,
        vybe_ast::ProtocolSlot::Repr,
    ] {
        chunk.emit_dup(line);
        chunk.emit_op_u16(Op::REF_FUNC, str_idx as u16, line);
        chunk.emit(0, line);
        let key = vybe_ast::protocol_slot_key(slot);
        let k = chunk.add_constant(Value::String(std::sync::Arc::from(key.as_str())));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
    }
}

/// `uuid.uuid4()` — RFC 4122 version 4, from `web:crypto.randomUUID`.
pub fn emit_uuid4(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    call_import(chunks, current, "web:crypto", "randomUUID", 0, line);
    wrap_uuid(chunks, current, line);
}

/// `uuid.UUID(hex)` — normalise the text CPython accepts (braces, urn prefix,
/// dashes anywhere) back to the canonical form.
pub fn emit_uuid_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let digits = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], base, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    for junk in ["urn:uuid:", "{", "}", "-"] {
        chunks[current].emit_string_const(junk, line);
        chunks[current].emit_string_const("", line);
        call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    }
    lset(&mut chunks[current], digits, line);

    // 8-4-4-4-12 from the 32 hex digits.
    let substring = chunks[current].add_import("ecma:string", "substring");
    let concat = chunks[current].add_import("wasm:js-string", "concat");
    let mut first = true;
    for (start, end) in [(0, 8), (8, 12), (12, 16), (16, 20), (20, 32)] {
        if !first {
            chunks[current].emit_string_const("-", line);
            chunks[current].emit_call(concat, 2, line);
        }
        lget(&mut chunks[current], digits, line);
        chunks[current].emit_i32_const(start, line);
        chunks[current].emit_i32_const(end, line);
        chunks[current].emit_call(substring, 3, line);
        if !first {
            chunks[current].emit_call(concat, 2, line);
        }
        first = false;
    }
    wrap_uuid(chunks, current, line);
}
