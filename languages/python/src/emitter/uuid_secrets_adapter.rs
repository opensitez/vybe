//! Python `uuid` and `secrets` — both are randomness surfaces, both sit on
//! `web:crypto`.
//!
//! `web:crypto.getRandomValues` fills a plain `ObjectKind::Array` (the host
//! leaves a TypedArray untouched), so the byte generators build an array,
//! fill it, and convert the result to the `Uint8Array` Python calls `bytes`.

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::{base64, collections};

use super::adapter_util::{call_import, lget, lset, new_object, stash_exact, struct_get, struct_set};

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
    if let Some(idx) = chunks.iter().position(|c| c.name == "__py_uuid_str") {
        return idx;
    }
    let mut helper = Chunk::new("__py_uuid_str");
    helper.arity = 1;
    helper.local_count = helper.local_count.max(1);
    helper.emit_op_u16(Op::LOCAL_GET, 0, line);
    let k = class_slots::resolve_interned(
        &mut helper,
        &ClassSlot::internal(CANONICAL_KEY),
        &PlainNames,
    );
    class_slots::emit_class_get(&mut helper, ObjSource::Stack, &k, Dest::Stack, line);
    helper.emit_op(Op::RETURN, line);
    chunks.push(helper);
    chunks.len() - 1
}

fn build_eq_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == "__py_uuid_eq") {
        return idx;
    }
    let mut helper = Chunk::new("__py_uuid_eq");
    helper.arity = 2;
    helper.local_count = helper.local_count.max(2);
    helper.emit_op_u16(Op::LOCAL_GET, 0, line);
    struct_get(&mut helper, &ClassSlot::internal(CANONICAL_KEY), line);
    helper.emit_op_u16(Op::LOCAL_GET, 1, line);
    struct_get(&mut helper, &ClassSlot::internal(CANONICAL_KEY), line);
    let eq = helper.add_import("wasm:js-string", "equals");
    helper.emit_call(eq, 2, line);
    helper.emit_op(Op::RETURN, line);
    chunks.push(helper);
    chunks.len() - 1
}

fn set_protocol_slot(chunk: &mut Chunk, idx: usize, slot: vybe_ast::ProtocolSlot, line: u32) {
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::REF_FUNC, idx as u16, line);
    chunk.emit(0, line);
    let key = vybe_ast::protocol_slot_key(slot);
    let cs_slot = class_slots::resolve(
        &ClassSlot::Internal((key.as_str()).to_string()),
        &PlainNames,
    );
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_slot, ValueSource::Stack, line);
}

fn set_named_slot(chunk: &mut Chunk, idx: usize, name: &str, line: u32) {
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::REF_FUNC, idx as u16, line);
    chunk.emit(0, line);
    let cs_slot = class_slots::resolve(&ClassSlot::internal(name), &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_slot, ValueSource::Stack, line);
}

fn push_uuid_hex(chunks: &mut [Chunk], current: usize, text: u16, line: u32) {
    lget(&mut chunks[current], text, line);
    chunks[current].emit_string_const("-", line);
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
}

fn push_uuid_version(chunks: &mut [Chunk], current: usize, text: u16, fallback: i32, line: u32) {
    lget(&mut chunks[current], text, line);
    chunks[current].emit_i32_const(14, line);
    chunks[current].emit_i32_const(15, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    call_import(chunks, current, "ecma:number", "parseInt", 1, line);
    let slot = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], slot, line);
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "ecma:number", "isNaN", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(fallback, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], slot, line);
    chunks[current].emit_end(line);
}

fn push_substring(chunk: &mut Chunk, slot: u16, start: i32, end: i32, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_i32_const(start, line);
    chunk.emit_i32_const(end, line);
    let substring = chunk.add_import("ecma:string", "substring");
    chunk.emit_call(substring, 3, line);
}

fn concat_top(chunk: &mut Chunk, line: u32) {
    let concat = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(concat, 2, line);
}

fn push_uuid_from_hex_digest(chunk: &mut Chunk, hex: u16, version: i32, line: u32) {
    push_substring(chunk, hex, 0, 8, line);
    chunk.emit_string_const("-", line);
    concat_top(chunk, line);
    push_substring(chunk, hex, 8, 12, line);
    concat_top(chunk, line);
    chunk.emit_string_const("-", line);
    concat_top(chunk, line);
    chunk.emit_string_const(&version.to_string(), line);
    concat_top(chunk, line);
    push_substring(chunk, hex, 13, 16, line);
    concat_top(chunk, line);
    chunk.emit_string_const("-", line);
    concat_top(chunk, line);
    chunk.emit_string_const("8", line);
    concat_top(chunk, line);
    push_substring(chunk, hex, 17, 20, line);
    concat_top(chunk, line);
    chunk.emit_string_const("-", line);
    concat_top(chunk, line);
    push_substring(chunk, hex, 20, 32, line);
    concat_top(chunk, line);
}

/// Build the `UUID` object around a canonical string already on the stack.
fn wrap_uuid_with_version(chunks: &mut Vec<Chunk>, current: usize, version: i32, line: u32) {
    let str_idx = build_str_helper(chunks, line);
    let eq_idx = build_eq_helper(chunks, line);
    let text = chunks[current].alloc_scratch(1);
    let hex = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], text, line);
    push_uuid_hex(chunks, current, text, line);
    lset(&mut chunks[current], hex, line);

    let chunk = &mut chunks[current];
    new_object(chunk, line);
    chunk.emit_dup(line);
    chunk.emit_string_const("UUID", line);
    let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_id, ValueSource::Stack, line);
    chunk.emit_dup(line);
    lget(chunk, text, line);
    struct_set(chunk, &ClassSlot::internal(CANONICAL_KEY), line);
    // `.hex` is the same digits without the dashes.
    chunk.emit_dup(line);
    lget(chunk, hex, line);
    struct_set(chunk, &ClassSlot::internal("hex"), line);
    chunk.emit_dup(line);
    lget(chunk, hex, line);
    let from_hex = chunk.add_import("ecma:uint8array", "fromHex");
    chunk.emit_call(from_hex, 1, line);
    struct_set(chunk, &ClassSlot::internal("bytes"), line);
    chunk.emit_dup(line);
    chunk.emit_i32_const(1, line);
    struct_set(chunk, &ClassSlot::internal("int"), line);
    chunk.emit_dup(line);
    chunk.emit_i32_const(0, line);
    struct_set(chunk, &ClassSlot::internal("time_low"), line);
    chunk.emit_dup(line);
    chunk.emit_i32_const(0, line);
    struct_set(chunk, &ClassSlot::internal("clock_seq"), line);
    chunk.emit_dup(line);
    chunk.emit_i32_const(version, line);
    struct_set(chunk, &ClassSlot::internal("version"), line);
    chunk.emit_dup(line);
    chunk.emit_string_const("specified in RFC 4122", line);
    struct_set(chunk, &ClassSlot::internal("variant"), line);
    chunk.emit_dup(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_i32_const(version, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_array_new_fixed(0, 6, line);
    struct_set(chunk, &ClassSlot::internal("fields"), line);
    chunk.emit_dup(line);
    chunk.emit_string_const("urn:uuid:", line);
    lget(chunk, text, line);
    let concat = chunk.add_import("wasm:js-string", "concat");
    // `concat(a, b)` — the prefix has to be the LEFT operand.
    chunk.emit_call(concat, 2, line);
    struct_set(chunk, &ClassSlot::internal("urn"), line);

    set_protocol_slot(chunk, str_idx, vybe_ast::ProtocolSlot::ToString, line);
    set_protocol_slot(chunk, str_idx, vybe_ast::ProtocolSlot::Repr, line);
    set_protocol_slot(chunk, eq_idx, vybe_ast::ProtocolSlot::Eq, line);
    set_named_slot(chunk, str_idx, "__str__", line);
    set_named_slot(chunk, str_idx, "__repr__", line);
    set_named_slot(chunk, eq_idx, "__eq__", line);
}

fn wrap_uuid(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    wrap_uuid_with_version(chunks, current, 0, line);
}

/// `uuid.uuid4()` — RFC 4122 version 4, from `web:crypto.randomUUID`.
pub fn emit_uuid4(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    call_import(chunks, current, "web:crypto", "randomUUID", 0, line);
    wrap_uuid_with_version(chunks, current, 4, line);
}

/// `uuid.UUID(hex)` — normalise the text CPython accepts (braces, urn prefix,
/// dashes anywhere) back to the canonical form.
pub fn emit_uuid_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let digits = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], base, line);
    call_import(chunks, current, "ecma:arraybuffer", "isView", 1, line);
    call_import(chunks, current, "wasm:js-boolean", "cast", 1, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], base, line);
    call_import(chunks, current, "ecma:uint8array", "toHex", 1, line);
    lset(&mut chunks[current], digits, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], base, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    for junk in ["urn:uuid:", "{", "}", "-"] {
        chunks[current].emit_string_const(junk, line);
        chunks[current].emit_string_const("", line);
        call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    }
    lset(&mut chunks[current], digits, line);
    chunks[current].emit_end(line);

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

fn emit_uuid_name_based(chunks: &mut Vec<Chunk>, current: usize, argc: u8, version: i32, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let h = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let hex = chunks[current].alloc_scratch(1);

    chunks[current].emit_string_const(if version == 3 { "md5" } else { "sha1" }, line);
    call_import(chunks, current, "node:crypto", "createHash", 1, line);
    lset(&mut chunks[current], h, line);

    chunks[current].emit_i32_const(version, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_string_const(":", line);
    concat_top(&mut chunks[current], line);
    lget(&mut chunks[current], base, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    concat_top(&mut chunks[current], line);
    chunks[current].emit_string_const(":", line);
    concat_top(&mut chunks[current], line);
    lget(&mut chunks[current], base + 1, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    concat_top(&mut chunks[current], line);
    lset(&mut chunks[current], data, line);

    lget(&mut chunks[current], h, line);
    lget(&mut chunks[current], data, line);
    call_import(chunks, current, "node:crypto", "_hashUpdate", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], h, line);
    chunks[current].emit_string_const("hex", line);
    call_import(chunks, current, "node:crypto", "_hashDigest", 2, line);
    lset(&mut chunks[current], hex, line);
    push_uuid_from_hex_digest(&mut chunks[current], hex, version, line);
    wrap_uuid_with_version(chunks, current, version, line);
}

pub fn emit_uuid3(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_uuid_name_based(chunks, current, argc, 3, line);
}

pub fn emit_uuid5(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_uuid_name_based(chunks, current, argc, 5, line);
}
