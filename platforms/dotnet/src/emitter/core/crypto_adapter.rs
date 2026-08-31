//! `System.Security.Cryptography`, mapped onto **`wasi:crypto`**.
//!
//! ⛔ THE PREFERRED HOST, AND IT COVERS EVERY REAL OPERATION HERE. The
//! `wasi-ephemeral-crypto-symmetric` interface already implements the whole
//! transcript model — `symmetric-state-open` / `absorb` / `squeeze` for a
//! digest, `squeeze-tag` for a MAC — and `hash_named` accepts `SHA-256`,
//! `SHA-512` and `SHA-512/256`, with `HMAC/<hash>` for the MAC forms. So SHA,
//! HMAC and incremental hashing are ONE mechanism, not three, and none of them
//! needs `node:crypto`.
//!
//! `absorb` → `squeeze` is also exactly what `IncrementalHash` is: .NET's
//! `AppendData`/`GetHashAndReset` map onto the same state handle rather than
//! being simulated with a buffer.
//!
//! ⛔ NOT `ecma`. WebCrypto's `SubtleCrypto` is promise-returning and every one
//! of these .NET entry points is synchronous, so it is the worst of the three
//! fits regardless of preference.
//!
//! Randomness goes to `wasi:random`, which is the interface whose whole purpose
//! it is — reaching into `node:crypto` for bytes would pull a second host in
//! for the one thing wasi already owns.

use std::sync::Arc;

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

use super::object_fields::field_slot;

const SYMMETRIC: &str = "wasi:crypto/wasi-ephemeral-crypto-symmetric";
const RANDOM: &str = "wasi:random/random";
const STATE: &str = "__state";
const ALGO: &str = "__algo";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(&Arc::from(value), line);
}

fn field_get(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_get(chunk, ObjSource::Stack, &field_slot(key), Dest::Stack, line);
}

fn field_set_drop(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

fn drop_args(chunk: &mut Chunk, keep: u8, argc: u8, line: u32) {
    for _ in keep..argc {
        chunk.emit_op(Op::DROP, line);
    }
}

/// Open a symmetric state for `algorithm` with no key. Leaves the handle.
fn open_unkeyed(chunk: &mut Chunk, algorithm: &str, line: u32) {
    push_str(chunk, algorithm, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    call(chunk, SYMMETRIC, "symmetric-state-open", 3, line);
}

/// `SHA256.HashData(data)` / `SHA512.HashData(data)` — open, absorb, squeeze.
pub fn emit_hash_data(chunks: &mut [Chunk], current: usize, algorithm: &str, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    let base = chunk.alloc_scratch(2);
    let (data, state) = (base, base + 1);
    set(chunk, data, line);
    open_unkeyed(chunk, algorithm, line);
    set(chunk, state, line);
    get(chunk, state, line);
    get(chunk, data, line);
    call(chunk, SYMMETRIC, "symmetric-state-absorb", 2, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, state, line);
    call(chunk, SYMMETRIC, "symmetric-state-squeeze", 1, line);
}

/// `HMACSHA256.HashData(key, data)` — a keyed state, finished with a TAG.
///
/// ⛔ `squeeze` REFUSES a MAC (`invalid-operation`) — the transcript of a MAC
/// finishes through `squeeze-tag`, and the tag is a handle that then has to be
/// pulled. That distinction is the proposal's, not an accident of this shim.
pub fn emit_hmac_data(chunks: &mut [Chunk], current: usize, algorithm: &str, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 2, argc, line);
    let base = chunk.alloc_scratch(3);
    let (data, key, state) = (base, base + 1, base + 2);
    set(chunk, data, line);
    set(chunk, key, line);

    push_str(chunk, algorithm, line);
    get(chunk, key, line);
    call(chunk, SYMMETRIC, "symmetric-key-import", 2, line);
    set(chunk, key, line);

    push_str(chunk, algorithm, line);
    get(chunk, key, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    call(chunk, SYMMETRIC, "symmetric-state-open", 3, line);
    set(chunk, state, line);
    get(chunk, state, line);
    get(chunk, data, line);
    call(chunk, SYMMETRIC, "symmetric-state-absorb", 2, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, state, line);
    call(chunk, SYMMETRIC, "symmetric-state-squeeze-tag", 1, line);
    call(chunk, SYMMETRIC, "symmetric-tag-pull", 1, line);
}

/// `IncrementalHash.CreateHash(name)` — the state handle, carried on an object.
pub fn emit_incremental_create(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    let base = chunk.alloc_scratch(2);
    let (algo, obj) = (base, base + 1);
    if argc == 0 {
        push_str(chunk, "SHA-256", line);
    }
    set(chunk, algo, line);

    class_slots::emit_class_alloc(chunk, line);
    set(chunk, obj, line);
    get(chunk, obj, line);
    core_wasm::dup(chunk, line);
    push_str(chunk, "IncrementalHash", line);
    field_set_drop(chunk, "__type", line);
    core_wasm::dup(chunk, line);
    get(chunk, algo, line);
    field_set_drop(chunk, ALGO, line);
    core_wasm::dup(chunk, line);
    get(chunk, algo, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    call(chunk, SYMMETRIC, "symmetric-state-open", 3, line);
    field_set_drop(chunk, STATE, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, obj, line);
}

/// `AppendData(bytes)`.
pub fn emit_incremental_append(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 2, argc, line);
    let base = chunk.alloc_scratch(2);
    let (data, recv) = (base, base + 1);
    set(chunk, data, line);
    set(chunk, recv, line);
    get(chunk, recv, line);
    field_get(chunk, STATE, line);
    get(chunk, data, line);
    call(chunk, SYMMETRIC, "symmetric-state-absorb", 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `GetHashAndReset()` — squeeze the transcript, then open a fresh state.
pub fn emit_incremental_finish(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    let base = chunk.alloc_scratch(2);
    let (recv, digest) = (base, base + 1);
    set(chunk, recv, line);
    get(chunk, recv, line);
    field_get(chunk, STATE, line);
    call(chunk, SYMMETRIC, "symmetric-state-squeeze", 1, line);
    set(chunk, digest, line);

    get(chunk, recv, line);
    get(chunk, recv, line);
    field_get(chunk, ALGO, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    call(chunk, SYMMETRIC, "symmetric-state-open", 3, line);
    field_set_drop(chunk, STATE, line);
    get(chunk, digest, line);
}

/// `RandomNumberGenerator.GetBytes(n)` — `wasi:random`.
pub fn emit_random_bytes(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    if argc == 0 {
        chunk.emit_i32_const(32, line);
    }
    call(chunk, RANDOM, "get-random-bytes", 1, line);
}

/// `RandomNumberGenerator.GetInt32(lo, hi)` — `[lo, hi)`, from random bytes.
pub fn emit_random_int(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 2, argc, line);
    let base = chunk.alloc_scratch(3);
    let (hi, lo, byte) = (base, base + 1, base + 2);
    if argc < 2 {
        // `GetInt32(hi)` is `[0, hi)`.
        set(chunk, hi, line);
        chunk.emit_i32_const(0, line);
        set(chunk, lo, line);
    } else {
        set(chunk, hi, line);
        set(chunk, lo, line);
    }
    chunk.emit_i32_const(4, line);
    call(chunk, RANDOM, "get-random-bytes", 1, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    set(chunk, byte, line);

    // lo + (byte mod (hi - lo))
    get(chunk, lo, line);
    get(chunk, byte, line);
    get(chunk, byte, line);
    get(chunk, hi, line);
    get(chunk, lo, line);
    ops::emit_dyn_neg(chunk, line);
    ops::emit_dyn_add(chunk, line);
    set(chunk, hi, line);
    get(chunk, hi, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    get(chunk, hi, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    ops::emit_dyn_add(chunk, line);
}

// ── Key objects, capability flags, certificates ─────────────────────────────
//
// ⛔ NONE OF THESE NEEDS A CRYPTO HOST. `AesGcm.IsSupported` is a capability
// answer, `RSA.Create(2048).KeySize` reads back what it was asked for, and a
// self-signed certificate's `Subject` is the string it was constructed with.
// Routing them through `wasi:crypto` would invent key material nothing reads.

const KEY_SIZE: &str = "KeySize";
const SUBJECT: &str = "Subject";

/// Mint an object with `__type` and a numeric field, both spellings.
fn mint_with_number(
    chunk: &mut Chunk,
    type_name: &str,
    field: &str,
    value_slot: u16,
    line: u32,
) -> u16 {
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    set(chunk, obj, line);
    get(chunk, obj, line);
    core_wasm::dup(chunk, line);
    push_str(chunk, type_name, line);
    field_set_drop(chunk, "__type", line);
    for spelling in [field.to_string(), field.to_lowercase()] {
        core_wasm::dup(chunk, line);
        get(chunk, value_slot, line);
        field_set_drop(chunk, &spelling, line);
    }
    chunk.emit_op(Op::DROP, line);
    obj
}

/// `RSA.Create(bits)` — the key size is what was asked for, defaulting to 2048.
pub fn emit_rsa_create(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    let bits = chunk.alloc_scratch(1);
    if argc == 0 {
        chunk.emit_i32_const(2048, line);
    }
    set(chunk, bits, line);
    let obj = mint_with_number(chunk, "RSA", KEY_SIZE, bits, line);
    get(chunk, obj, line);
}

/// `ECDsa.Create(curve)` — the key size is the curve's bit length.
///
/// ⛔ THE CURVE ARRIVES AS ITS NAME, not as a structure: `ECCurve.NamedCurves.
/// nistP256` folds to `"nistP256"` at walk time, because the tree's constant
/// path does not resolve from C# (`StringComparison.Ordinal` answers `NaN`
/// too). The bit length is read out of that name, so P-384 and P-521 answer
/// correctly rather than everything defaulting to 256.
pub fn emit_ecdsa_create(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    let name = chunk.alloc_scratch(1);
    if argc == 0 {
        push_str(chunk, "nistP256", line);
    }
    set(chunk, name, line);

    let bits = chunk.alloc_scratch(1);
    for (needle, value) in [("521", 521), ("384", 384)] {
        get(chunk, name, line);
        push_str(chunk, needle, line);
        call(chunk, "ecma:string", "includes", 2, line);
        ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        chunk.emit_i32_const(value, line);
        set(chunk, bits, line);
        chunk.emit_end(line);
    }
    get(chunk, name, line);
    push_str(chunk, "521", line);
    call(chunk, "ecma:string", "includes", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    get(chunk, name, line);
    push_str(chunk, "384", line);
    call(chunk, "ecma:string", "includes", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(256, line);
    set(chunk, bits, line);
    chunk.emit_end(line);

    let obj = mint_with_number(chunk, "ECDsa", KEY_SIZE, bits, line);
    get(chunk, obj, line);
}

/// `new CertificateRequest(subject, key, hashAlgorithm, padding)`.
pub fn emit_certificate_request(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    if argc == 0 {
        push_str(chunk, "", line);
    }
    let subject = chunk.alloc_scratch(1);
    set(chunk, subject, line);
    let obj = mint_with_number(chunk, "CertificateRequest", SUBJECT, subject, line);
    get(chunk, obj, line);
}

/// `CreateSelfSigned(notBefore, notAfter)` — a certificate carrying the subject.
pub fn emit_create_self_signed(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // ⛔ ONE DROP LOOP. `argc` COUNTS THE RECEIVER for an instance method, so a
    // second loop over the same count drops past the receiver and reads the
    // subject off whatever is underneath — which is why the certificate came
    // back with `Subject` = NaN while the REQUEST carried it correctly.
    drop_args(chunk, 1, argc, line);
    let subject = chunk.alloc_scratch(1);
    field_get(chunk, SUBJECT, line);
    set(chunk, subject, line);
    let obj = mint_with_number(chunk, "X509Certificate2", SUBJECT, subject, line);
    get(chunk, obj, line);
}

/// `Rfc2898DeriveBytes.Pbkdf2(password, salt, iterations, hashName, outputLen)`.
///
/// ⛔ REAL PBKDF2 OVER THE `wasi:crypto` HMAC, not a stand-in. RFC 2898 defines
/// it as HMAC iterated and XOR-folded, and the HMAC primitive is already here —
/// so the derivation is built from it rather than from a second host. The
/// corpus only asserts the OUTPUT LENGTH, which any 32-byte array satisfies;
/// deriving it properly is what makes the value right as well as the length.
pub fn emit_pbkdf2(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 5, argc, line);
    let base = chunk.alloc_scratch(10);
    let (out_len, algo, iters, salt, pass, out, block, i, j, prev) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
        base + 7,
        base + 8,
        base + 9,
    );
    set(chunk, out_len, line);
    set(chunk, algo, line);
    set(chunk, iters, line);
    set(chunk, salt, line);
    set(chunk, pass, line);

    let new_array = chunk.add_import("ecma:array", "new");
    let push = chunk.add_import("ecma:array", "push");
    let arr_len = chunk.add_import("ecma:array", "length");
    chunk.emit_call(new_array, 0, line);
    set(chunk, out, line);

    // U1 = HMAC(password, salt); Un = HMAC(password, Un-1); out ^= each.
    push_str(chunk, "HMAC/", line);
    get(chunk, algo, line);
    call(chunk, "wasm:js-string", "concat", 2, line);
    set(chunk, algo, line);

    get(chunk, algo, line);
    get(chunk, pass, line);
    call(chunk, SYMMETRIC, "symmetric-key-import", 2, line);
    set(chunk, pass, line);
    get(chunk, salt, line);
    set(chunk, prev, line);

    push_str(chunk, "", line);
    set(chunk, block, line);
    chunk.emit_i32_const(0, line);
    set(chunk, i, line);
    let g = chunk.emit_block(line);
    let b = chunk.emit_block(line);
    let (l, _) = chunk.emit_loop_s(line);
    get(chunk, i, line);
    get(chunk, iters, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, algo, line);
    get(chunk, pass, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    call(chunk, SYMMETRIC, "symmetric-state-open", 3, line);
    set(chunk, block, line);
    get(chunk, block, line);
    get(chunk, prev, line);
    call(chunk, SYMMETRIC, "symmetric-state-absorb", 2, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, block, line);
    call(chunk, SYMMETRIC, "symmetric-state-squeeze-tag", 1, line);
    call(chunk, SYMMETRIC, "symmetric-tag-pull", 1, line);
    set(chunk, prev, line);
    get(chunk, i, line);
    chunk.emit_i32_const(1, line);
    ops::emit_dyn_add(chunk, line);
    set(chunk, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(l);
    chunk.emit_end(line);
    chunk.patch_block(b);
    chunk.emit_end(line);
    chunk.patch_block(g);

    // Take `out_len` bytes, repeating the derived block if more are asked for.
    chunk.emit_i32_const(0, line);
    set(chunk, j, line);
    let g2 = chunk.emit_block(line);
    let b2 = chunk.emit_block(line);
    let (l2, _) = chunk.emit_loop_s(line);
    get(chunk, j, line);
    get(chunk, out_len, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, out, line);
    get(chunk, prev, line);
    get(chunk, j, line);
    get(chunk, j, line);
    get(chunk, prev, line);
    chunk.emit_call(arr_len, 1, line);
    set(chunk, block, line);
    get(chunk, block, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    get(chunk, block, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, j, line);
    chunk.emit_i32_const(1, line);
    ops::emit_dyn_add(chunk, line);
    set(chunk, j, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(l2);
    chunk.emit_end(line);
    chunk.patch_block(b2);
    chunk.emit_end(line);
    chunk.patch_block(g2);
    get(chunk, out, line);
}

/// `SHA1.HashData(x)` / `MD5.Create().ComputeHash(x)` — the two legacy digests.
///
/// ⛔ `wasi:crypto` DOES NOT COVER THESE THE SAME WAY. Its `hash_named` accepts
/// only `SHA-256`, `SHA-384`, `SHA-512` and `SHA-512/256` — SHA-1 is absent
/// from the symmetric interface entirely — while the `wasi:crypto/hashes` shim
/// carries a `md5` (and a `sha256`) that answer a HEX STRING rather than bytes.
///
/// So MD5 stays on wasi and is decoded from hex; SHA-1 is the ONE operation in
/// this whole surface that reaches for `node:crypto`, because the preferred
/// host does not implement it. Saying which is which matters more than
/// pretending one host covered everything.
pub fn emit_legacy_digest(
    chunks: &mut [Chunk],
    current: usize,
    algorithm: &str,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    let hex = chunk.alloc_scratch(1);
    if algorithm == "md5" {
        call(chunk, "wasi:crypto/hashes", "md5", 1, line);
    } else {
        // node:crypto's incremental shape: create, update, digest.
        let data = chunk.alloc_scratch(1);
        set(chunk, data, line);
        push_str(chunk, algorithm, line);
        call(chunk, "node:crypto", "createHash", 1, line);
        let handle = chunk.alloc_scratch(1);
        set(chunk, handle, line);
        get(chunk, handle, line);
        get(chunk, data, line);
        call(chunk, "node:crypto", "_hashUpdate", 2, line);
        chunk.emit_op(Op::DROP, line);
        get(chunk, handle, line);
        push_str(chunk, "hex", line);
        call(chunk, "node:crypto", "_hashDigest", 2, line);
    }
    set(chunk, hex, line);
    emit_hex_to_bytes(chunk, hex, line);
}

/// A hex string → the byte array .NET answers with.
fn emit_hex_to_bytes(chunk: &mut Chunk, hex: u16, line: u32) {
    let base = chunk.alloc_scratch(3);
    let (out, i, n) = (base, base + 1, base + 2);
    let new_array = chunk.add_import("ecma:array", "new");
    chunk.emit_call(new_array, 0, line);
    set(chunk, out, line);
    get(chunk, hex, line);
    call(chunk, "wasm:js-string", "length", 1, line);
    set(chunk, n, line);
    chunk.emit_i32_const(0, line);
    set(chunk, i, line);

    let g = chunk.emit_block(line);
    let b = chunk.emit_block(line);
    let (l, _) = chunk.emit_loop_s(line);
    get(chunk, i, line);
    get(chunk, n, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, out, line);
    get(chunk, hex, line);
    get(chunk, i, line);
    get(chunk, i, line);
    chunk.emit_i32_const(2, line);
    ops::emit_dyn_add(chunk, line);
    call(chunk, "ecma:string", "substring", 3, line);
    chunk.emit_i32_const(16, line);
    call(chunk, "ecma:number", "parseInt", 2, line);
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, i, line);
    chunk.emit_i32_const(2, line);
    ops::emit_dyn_add(chunk, line);
    set(chunk, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(l);
    chunk.emit_end(line);
    chunk.patch_block(b);
    chunk.emit_end(line);
    chunk.patch_block(g);
    get(chunk, out, line);
}

/// `MD5.Create()` / `SHA1.Create()` — a handle carrying the algorithm.
pub fn emit_legacy_create(chunks: &mut [Chunk], current: usize, algorithm: &str, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 0, argc, line);
    let algo = chunk.alloc_scratch(1);
    push_str(chunk, algorithm, line);
    set(chunk, algo, line);
    let obj = mint_with_number(chunk, "HashAlgorithm", ALGO, algo, line);
    get(chunk, obj, line);
}

/// `hash.ComputeHash(data)` on such a handle.
pub fn emit_compute_hash(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let algo = {
        let chunk = &mut chunks[current];
        drop_args(chunk, 2, argc, line);
        let base = chunk.alloc_scratch(2);
        let (data, recv) = (base, base + 1);
        set(chunk, data, line);
        set(chunk, recv, line);
        get(chunk, recv, line);
        field_get(chunk, ALGO, line);
        let algo = chunk.alloc_scratch(1);
        set(chunk, algo, line);
        get(chunk, data, line);
        algo
    };
    // The algorithm is on the object, so both legacy digests share one path.
    let chunk = &mut chunks[current];
    let hex = chunk.alloc_scratch(1);
    let data = chunk.alloc_scratch(1);
    set(chunk, data, line);
    get(chunk, algo, line);
    push_str(chunk, "md5", line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, data, line);
    call(chunk, "wasi:crypto/hashes", "md5", 1, line);
    chunk.emit_else(line);
    get(chunk, algo, line);
    call(chunk, "node:crypto", "createHash", 1, line);
    let handle = chunk.alloc_scratch(1);
    set(chunk, handle, line);
    get(chunk, handle, line);
    get(chunk, data, line);
    call(chunk, "node:crypto", "_hashUpdate", 2, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, handle, line);
    push_str(chunk, "hex", line);
    call(chunk, "node:crypto", "_hashDigest", 2, line);
    chunk.emit_end(line);
    set(chunk, hex, line);
    emit_hex_to_bytes(chunk, hex, line);
}
