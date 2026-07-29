//! Python `hashlib` / `hmac` over `node:crypto`'s streaming digest objects.
//!
//! `createHash(algo)` / `createHmac(algo, key)` return host objects that hold
//! the digest state; `_hashUpdate`/`_hashDigest`/`_hashCopy` (and the hmac
//! pair) operate on them. Those host fns take the receiver as their FIRST
//! argument — the `update`/`digest` properties the host stamps on the object
//! are bare `HostFunction` refs without ecma's `__vybe_method_receiver`
//! marker, so `h.update(x)` would never pass `h`. This adapter calls the
//! receiver-taking fns directly with an explicit receiver instead.
//!
//! Python hands bytes as a `Uint8Array`, which the host's `bytes_from_value`
//! does not accept (it takes a string or an array of octets), so byte input is
//! widened through `ecma:array.from` first.
//!
//! Routed via `common:python.hash_*` from the profile; there is no prelude.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::errors::{emit_exception_new_finalize, emit_throw};
use vybe_compiler::primitives::ops;

const CRYPTO: &str = "node:crypto";

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    // Register on the CURRENT chunk so `normalize_import_table` remaps this
    // CALL_IMPORT through the emitting chunk's own table.
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

/// `obj[key] = <value on stack>`. Stack: `[obj, value] -> []`.
fn struct_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn struct_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
}

/// Stash `argc` call arguments into consecutive scratch slots, arg0 at `base`.
fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

/// `(digest_size, block_size)` per CPython's `hashlib`, for the algorithms
/// `node:crypto` implements. The host accepts both OpenSSL spelling
/// (`sha3-256`) and hashlib's (`sha3_256`), so the Python name passes
/// straight through.
///
/// `None` means unbacked: the adapter raises `ValueError` rather than let the
/// host return something plausible-looking. `shake_*` is excluded on purpose —
/// CPython's SHAKE takes a length argument (`h.hexdigest(20)`), which this
/// fixed-output protocol cannot express.
fn algo_sizes(algo: &str) -> Option<(i32, i32)> {
    Some(match algo {
        "md5" => (16, 64),
        "sha1" => (20, 64),
        "sha224" => (28, 64),
        "sha256" => (32, 64),
        "sha384" => (48, 128),
        "sha512" => (64, 128),
        "sha3_224" => (28, 144),
        "sha3_256" => (32, 136),
        "sha3_384" => (48, 104),
        "sha3_512" => (64, 72),
        "blake2b" => (64, 128),
        "blake2s" => (32, 64),
        _ => return None,
    })
}

/// Widen a data argument in `slot` to something `bytes_from_value` accepts:
/// a string passes through, bytes/list go through `ecma:array.from` (which
/// yields an array of octets). Leaves the widened value on the stack.
fn push_widened(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    // typeof(v) == "string" ? v : ecma:array.from(v)
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], slot, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "ecma:array", "from", 1, line);
    chunks[current].emit_end(line);
}

/// Stamp the Python-visible attributes onto a host digest object held in
/// `slot`: `name`, `digest_size`, `block_size`.
fn stamp_attrs(chunks: &mut [Chunk], current: usize, slot: u16, name: &str, algo: &str, line: u32) {
    let (digest_size, block_size) = algo_sizes(algo).unwrap_or((32, 64));
    lget(&mut chunks[current], slot, line);
    chunks[current].emit_string_const(name, line);
    struct_set_key(&mut chunks[current], "name", line);
    lget(&mut chunks[current], slot, line);
    chunks[current].emit_i32_const(digest_size, line);
    struct_set_key(&mut chunks[current], "digest_size", line);
    lget(&mut chunks[current], slot, line);
    chunks[current].emit_i32_const(block_size, line);
    struct_set_key(&mut chunks[current], "block_size", line);
}

/// Throw `ValueError: unsupported hash type <algo>`.
fn throw_unsupported(chunks: &mut [Chunk], current: usize, algo: &str, line: u32) {
    chunks[current].emit_string_const(&format!("unsupported hash type {algo}"), line);
    emit_exception_new_finalize(&mut chunks[current], "ValueError", line);
    emit_throw(&mut chunks[current], line);
}

/// `hashlib.<algo>([data])` / `hashlib.new(algo[, data])` for a statically
/// known algorithm. Leaves the digest object on the stack.
fn emit_new_for_algo(chunks: &mut [Chunk], current: usize, algo: &str, argc: u8, line: u32) {
    if algo_sizes(algo).is_none() {
        // Drop the arguments, then raise — the value never materialises.
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        throw_unsupported(chunks, current, algo, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let h = chunks[current].alloc_scratch(1);

    chunks[current].emit_string_const(algo, line);
    call_import(chunks, current, CRYPTO, "createHash", 1, line);
    lset(&mut chunks[current], h, line);

    if argc >= 1 {
        lget(&mut chunks[current], h, line);
        push_widened(chunks, current, base, line);
        call_import(chunks, current, CRYPTO, "_hashUpdate", 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    stamp_attrs(chunks, current, h, algo, algo, line);
    lget(&mut chunks[current], h, line);
}

/// `hashlib.sha256(...)` and friends — the algorithm is in the builtin name.
pub fn emit_sha256(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_new_for_algo(chunks, current, "sha256", argc, line);
}

pub fn emit_sha512(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_new_for_algo(chunks, current, "sha512", argc, line);
}

pub fn emit_sha1(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_new_for_algo(chunks, current, "sha1", argc, line);
}

pub fn emit_md5(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_new_for_algo(chunks, current, "md5", argc, line);
}

/// `hashlib.new(algo[, data])` — the algorithm is a runtime string, so the
/// host does the lookup and the size attributes come from a runtime chain.
pub fn emit_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let h = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], base, line);
    call_import(chunks, current, CRYPTO, "createHash", 1, line);
    lset(&mut chunks[current], h, line);

    if argc >= 2 {
        lget(&mut chunks[current], h, line);
        push_widened(chunks, current, base + 1, line);
        call_import(chunks, current, CRYPTO, "_hashUpdate", 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }

    // name = <the algo argument>; sizes resolve from it at runtime.
    lget(&mut chunks[current], h, line);
    lget(&mut chunks[current], base, line);
    struct_set_key(&mut chunks[current], "name", line);
    // Sizes resolve from the runtime algorithm string. Same table as
    // `algo_sizes`; an unknown name lands on the sha256 defaults, matching
    // what the host will actually compute for it.
    for (key, sizes, dflt) in [
        (
            "digest_size",
            [
                ("md5", 16), ("sha1", 20), ("sha224", 28), ("sha384", 48),
                ("sha512", 64), ("sha3_224", 28), ("sha3_384", 48),
                ("sha3_512", 64), ("blake2b", 64),
            ],
            32,
        ),
        (
            "block_size",
            [
                ("md5", 64), ("sha1", 64), ("sha224", 64), ("sha384", 128),
                ("sha512", 128), ("sha3_224", 144), ("sha3_384", 104),
                ("sha3_512", 72), ("blake2b", 128),
            ],
            64,
        ),
    ] {
        lget(&mut chunks[current], h, line);
        for (algo, size) in sizes {
            lget(&mut chunks[current], base, line);
            chunks[current].emit_string_const(algo, line);
            ops::emit_dyn_eq(&mut chunks[current], line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_i32_const(size, line);
            chunks[current].emit_else(line);
        }
        chunks[current].emit_i32_const(dflt, line);
        for _ in 0..sizes.len() {
            chunks[current].emit_end(line);
        }
        struct_set_key(&mut chunks[current], key, line);
    }
    lget(&mut chunks[current], h, line);
}

/// Push i32 1 when `slot` holds a `node:crypto` digest object — only
/// `createHash`/`createHmac` stamp `__algo`. Lets the shared `update`/`copy`
/// value-method adapters branch to the digest path without a name check.
pub fn emit_is_digest(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    struct_get_key(&mut chunks[current], "__algo", line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

/// `h.update(data)` with the receiver in `recv` and the data in `src`.
/// Leaves `None` on the stack (Python's `update` returns nothing).
pub fn emit_update_slots(chunks: &mut [Chunk], current: usize, recv: u16, src: u16, line: u32) {
    lget(&mut chunks[current], recv, line);
    push_widened(chunks, current, src, line);
    hmac_or_hash(chunks, current, recv, "_hmacUpdate", "_hashUpdate", line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `h.copy()` with the receiver in `recv` — an independent digest object
/// carrying the same Python-visible attributes.
pub fn emit_copy_slot(chunks: &mut [Chunk], current: usize, recv: u16, line: u32) {
    let c = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], recv, line);
    call_import(chunks, current, CRYPTO, "_hashCopy", 1, line);
    lset(&mut chunks[current], c, line);
    for key in ["name", "digest_size", "block_size"] {
        lget(&mut chunks[current], c, line);
        lget(&mut chunks[current], recv, line);
        struct_get_key(&mut chunks[current], key, line);
        struct_set_key(&mut chunks[current], key, line);
    }
    lget(&mut chunks[current], c, line);
}

/// Call the hmac or hash variant of a receiver-taking host fn depending on
/// whether the receiver carries `__key` (only `createHmac` stamps that).
/// Stack in: `[receiver, arg]`; stack out: the host fn's result.
fn hmac_or_hash(
    chunks: &mut [Chunk],
    current: usize,
    recv: u16,
    hmac_name: &str,
    hash_name: &str,
    line: u32,
) {
    lget(&mut chunks[current], recv, line);
    struct_get_key(&mut chunks[current], "__key", line);
    chunks[current].emit_op(Op::NULL, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    call_import(chunks, current, CRYPTO, hash_name, 2, line);
    chunks[current].emit_else(line);
    call_import(chunks, current, CRYPTO, hmac_name, 2, line);
    chunks[current].emit_end(line);
}

/// `h.hexdigest()` — lowercase hex string.
pub fn emit_hexdigest(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    lget(&mut chunks[current], base, line);
    chunks[current].emit_string_const("hex", line);
    hmac_or_hash(chunks, current, base, "_hmacDigest", "_hashDigest", line);
}

/// `h.digest()` — the same digest as raw bytes, via `bytes.fromhex`.
pub fn emit_digest(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_hexdigest(chunks, current, argc, line);
    call_import(chunks, current, "ecma:uint8array", "fromHex", 1, line);
}

/// `hmac.new(key[, msg[, digestmod]])`.
pub fn emit_hmac_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let h = chunks[current].alloc_scratch(1);

    if argc >= 3 {
        lget(&mut chunks[current], base + 2, line);
    } else {
        chunks[current].emit_string_const("sha256", line);
    }
    push_widened(chunks, current, base, line);
    call_import(chunks, current, CRYPTO, "createHmac", 2, line);
    lset(&mut chunks[current], h, line);

    // `createHmac` already stored the widened key under `__key`; that
    // property is also what marks the object as an hmac for `hmac_or_hash`.
    // Do NOT restamp it — the raw Uint8Array does not survive the host's
    // `bytes_from_value` and the digest would be computed with an empty key.

    if argc >= 2 {
        lget(&mut chunks[current], h, line);
        push_widened(chunks, current, base + 1, line);
        call_import(chunks, current, CRYPTO, "_hmacUpdate", 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    stamp_attrs(chunks, current, h, "hmac-sha256", "sha256", line);
    lget(&mut chunks[current], h, line);
}

/// `hmac.compare_digest(a, b)` — constant-time in CPython; equality here.
pub fn emit_compare_digest(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    lget(&mut chunks[current], base, line);
    lget(&mut chunks[current], base + 1, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}
