//! Getting guest data INTO and OUT OF linear memory, for the canonical ABI.
//!
//! The Component Model moves values through **linear memory**: `canon
//! stream.write` is `(handle, ptr, n)` and reads `n` elements starting at
//! `ptr`. Every other emitter in this compiler works on the GC/Value model —
//! a string is a `Value::String`, an array is a heap object — and linear
//! memory was previously touched in one place only, `channels.rs`, for futex
//! WORDS. There was no way to put a string's bytes where a conforming
//! component could read them.
//!
//! **This splices instructions at the CALL SITE** — authoring way (1), the
//! default. Deliberately NOT a `__stdlib_*` linked helper: that mechanism is
//! being retired, and a short encode-and-copy is none of the three things that
//! justify a linked function (a large body, recursion, or needing to be a
//! first-class value). Routing it through the helper bundle instead cost four
//! rounds against `MAPPINGS`, export ordering, `referenced_helper_exports` and
//! the dependency graph — and left python and php printing nothing at all,
//! because a helper that another helper calls is invisible to the scan that
//! decides which helpers get linked.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Where the bump pointer for canonical buffers lives. Shared with the channel
/// futex allocator on purpose: two bump pointers into one page would hand out
/// the same address twice.
const BUMP: &str = "__vybe_chan_futex_next";

/// Marshal a string into linear memory as UTF-8.
///
/// `push_string` leaves the string on the stack. Returns `(ptr_slot, len_slot)`
/// — scratch locals holding the address and the BYTE count.
///
/// Stack effect: `[] → []`; both results land in the returned slots.
pub fn emit_store_utf8(
    chunk: &mut Chunk,
    line: u32,
    push_string: impl FnOnce(&mut Chunk),
) -> (u16, u16) {
    emit_store_byte_array(chunk, line, |chunk| {
        // bytes = Array.from(TextEncoder().encode(s))
        //
        // ⚠ The `Array.from` is LOAD-BEARING. `web:encoding.encode` answers a
        // `Uint8Array` — `ObjectKind::TypedArray` — and nothing in reach indexes
        // one: `Op::ARRAY_GET` matches only `ObjectKind::Array`, and neither
        // `ecma:array.at` nor plain indexing handles the typed case. Without the
        // conversion every byte read as 0, so the write emitted the right COUNT of
        // NUL bytes and stdout looked simply EMPTY. `ecma:array.from` is the same
        // conversion `strings::emit_scalar_chars` already uses.
        let encoder_new = chunk.add_import("web:encoding", "encoderNew");
        chunk.emit_call(encoder_new, 0, line);
        push_string(chunk);
        let encode = chunk.add_import("web:encoding", "encode");
        chunk.emit_call(encode, 2, line);
        let from = chunk.add_import("ecma:array", "from");
        chunk.emit_call(from, 1, line);
    })
}

/// Marshal an ALREADY-BYTE-VALUED array into linear memory.
///
/// `push_array` leaves a plain `Array` of byte-valued numbers on the stack.
/// Returns `(ptr_slot, len_slot)` exactly as [`emit_store_utf8`] does — which is
/// now this function plus one `TextEncoder` pass in front.
///
/// ⚠**Binary data must come through HERE, never through `emit_store_utf8`.**
/// A byte array handed to the string form gets `TextEncoder`'d, i.e. its
/// DECIMAL RENDERING (`"72,101,108"`) is what reaches the file. That is silent:
/// the write succeeds, the byte count is plausible, and only the contents are
/// wrong. `copy` is the caller that forced this split — a copied PNG has no
/// business round-tripping through UTF-8.
pub fn emit_store_byte_array(
    chunk: &mut Chunk,
    line: u32,
    push_array: impl FnOnce(&mut Chunk),
) -> (u16, u16) {
    let bytes = chunk.alloc_scratch(1);
    let len = chunk.alloc_scratch(1);
    let ptr = chunk.alloc_scratch(1);
    let i = chunk.alloc_scratch(1);
    let base = chunk.alloc_scratch(1);

    push_array(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, bytes, line);

    // len = bytes.length — the UTF-8 BYTE count, not the string's `.length`,
    // which counts UTF-16 code units.
    chunk.emit_op_u16(Op::LOCAL_GET, bytes, line);
    let arr_len = chunk.add_import("ecma:array", "length");
    chunk.emit_call(arr_len, 1, line);
    let to_i32 = chunk.add_import("wasm:js-number", "toI32");
    chunk.emit_call(to_i32, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len, line);

    // ── bump-allocate `len` bytes ───────────────────────────────────────────
    // base = BUMP, growing a page the first time (BUMP starts null/undefined).
    crate::primitives::globals::emit_read(chunk, BUMP, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, base, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    let undef_test = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef_test, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::MEMORY_GROW, 0, line);
    chunk.emit_i32_const(65536, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base, line);
    chunk.emit_end(line);

    // Grow again if this request would run off the end of the page. Without
    // it the allocator hands out addresses past the memory it owns, and the
    // corruption surfaces nowhere near here.
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::MEMORY_SIZE, 0, line);
    chunk.emit_i32_const(65536, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op(Op::I32_GT_U, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::MEMORY_GROW, 0, line);
    chunk.emit_i32_const(65536, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ptr, line);

    // BUMP = base + align_to(len + 8, 8) — 8-aligned so a later i64/f64 store
    // lands legally, and never zero-width so two empty buffers cannot alias.
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    chunk.emit_i32_const(15, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_i32_const(!7, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_ADD, line);
    crate::primitives::globals::emit_write(chunk, BUMP, line);

    // ── copy ────────────────────────────────────────────────────────────────
    // `block { loop { br_if 1 (done); store; br 0 } }` — depth 1 leaves the
    // block, depth 0 re-enters the loop. `I32_STORE8` carries NO memarg: the
    // VM's is marker-tagged and optional, absent meaning natural align,
    // offset 0, memory 0.
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    let _done = chunk.emit_block(line);
    let (_copy, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let to_i32b = chunk.add_import("wasm:js-number", "toI32");
    chunk.emit_call(to_i32b, 1, line);
    chunk.emit_op(Op::I32_STORE8, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    (ptr, len)
}

/// Bump-allocate `n` bytes of linear memory. Returns a scratch local holding
/// the address.
///
/// Split out of [`emit_store_utf8`] because the READ direction needs the same
/// allocator and must not get a second one: two bump pointers into one page
/// hand out the same address twice, which is the exact failure the `BUMP`
/// constant's comment already warns about for the futex allocator.
///
/// `n_slot` holds the byte count. Stack effect: `[] → []`.
pub fn emit_alloc(chunk: &mut Chunk, line: u32, n_slot: u16) -> u16 {
    let ptr = chunk.alloc_scratch(1);
    let base = chunk.alloc_scratch(1);

    crate::primitives::globals::emit_read(chunk, BUMP, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, base, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    let undef_test = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef_test, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::MEMORY_GROW, 0, line);
    chunk.emit_i32_const(65536, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base, line);
    chunk.emit_end(line);

    // Grow while the request would run off the end. A LOOP, not a single `if`:
    // `emit_store_utf8`'s one-shot grow adds exactly one page, which is not
    // enough for a request larger than 64KB. A read buffer is caller-sized, so
    // this one has to keep growing until it fits.
    let done = chunk.emit_block(line);
    let (grow, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n_slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::MEMORY_SIZE, 0, line);
    chunk.emit_i32_const(65536, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_br_if(1, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::MEMORY_GROW, 0, line);
    chunk.emit_i32_const(65536, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(grow);
    chunk.emit_end(line);
    chunk.patch_block(done);

    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ptr, line);

    // BUMP = base + align_to(n + 8, 8), same rule as the store direction:
    // 8-aligned so a later i64/f64 store lands legally, never zero-width so
    // two empty buffers cannot alias.
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n_slot, line);
    chunk.emit_i32_const(15, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_i32_const(!7, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_ADD, line);
    crate::primitives::globals::emit_write(chunk, BUMP, line);

    ptr
}

/// A fresh, empty byte array. Stack effect: `[] → [array]`.
///
/// `newWithLength(0)`, then filled with [`emit_append_bytes`]. The store
/// direction indexes an `Array.from`ed encode result; the load direction owns
/// its array from the start, so `ecma:array.push` is the whole conversion.
pub fn emit_new_bytes(chunk: &mut Chunk, line: u32) {
    chunk.emit_i32_const(0, line);
    let new_arr = chunk.add_import("vybe:js-array", "newWithLength");
    chunk.emit_call(new_arr, 1, line);
}

/// Append `len_slot` bytes at `ptr_slot` to the byte array in `arr_slot`.
///
/// Stack effect: `[] → []` — the array is read from and written back to its
/// local, so a drain loop can call this once per chunk and still end up with
/// ONE array. That is the point: a decoder must see the whole byte run at
/// once. Decoding each chunk separately splits any multi-byte sequence that
/// straddles a chunk boundary into two replacement characters, and with a
/// fixed 64KB read buffer the boundary lands mid-sequence for roughly one
/// body in 65536 — a corruption that would pass every test with short input.
pub fn emit_append_bytes(chunk: &mut Chunk, line: u32, arr_slot: u16, ptr_slot: u16, len_slot: u16) {
    let i = chunk.alloc_scratch(1);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    let done = chunk.emit_block(line);
    let (copy, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::I32_ADD, line);
    // `I32_LOAD8_U` — UNSIGNED. The signed form sign-extends any byte ≥ 0x80,
    // so every multi-byte UTF-8 sequence would decode as a negative number and
    // the text would come back mojibake.
    chunk.emit_op(Op::I32_LOAD8_U, line);
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line); // push answers the new length

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(copy);
    chunk.emit_end(line);
    chunk.patch_block(done);
}

/// Decode a byte array as UTF-8. Stack effect: `[array] → [string]`.
///
/// `TextDecoder().decode(bytes)` — `decode(decoder, input)`, the same
/// two-argument shape `web:encoding` registers.
///
/// ⚠The array is converted to a **Uint8Array** first, and that is not
/// ceremony. `TextDecoder.decode` takes a `BufferSource` (WHATWG Encoding
/// §8.2) — an ArrayBuffer or a view over one — and the host follows the spec:
/// `bytes_from_arg` matches `ArrayBuffer`, `TypedArray`, or an object with a
/// `buffer`, and answers `Vec::new()` for anything else. A plain array of
/// numbers therefore decoded to `""` with no error raised, which is how this
/// shipped: every `emit_load_utf8` caller read an empty string and had no way
/// to tell that from an empty stream.
pub fn emit_decode_utf8(chunk: &mut Chunk, line: u32) {
    let bytes = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, bytes, line);
    let decoder_new = chunk.add_import("web:encoding", "decoderNew");
    chunk.emit_call(decoder_new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes, line);
    let to_u8 = chunk.add_import("ecma:uint8array", "newFromIterable");
    chunk.emit_call(to_u8, 1, line);
    let decode = chunk.add_import("web:encoding", "decode");
    chunk.emit_call(decode, 2, line);
}

/// Decode `len_slot` bytes of UTF-8 at `ptr_slot` back into a string.
///
/// The mirror of [`emit_store_utf8`], and needed for the same reason: `canon
/// stream.read` copies elements INTO LINEAR MEMORY and answers only a count,
/// so a guest that wants the text has to go and get it. Nothing else in this
/// compiler reads bytes back out — the store direction was written first
/// because output came first.
///
/// The one-shot form: for a run that arrives in pieces, use
/// [`emit_new_bytes`] + [`emit_append_bytes`] + [`emit_decode_utf8`] so the
/// decode sees all the bytes together.
///
/// Stack effect: `[] → [string]`.
pub fn emit_load_utf8(chunk: &mut Chunk, line: u32, ptr_slot: u16, len_slot: u16) {
    let bytes = chunk.alloc_scratch(1);
    emit_new_bytes(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, bytes, line);
    emit_append_bytes(chunk, line, bytes, ptr_slot, len_slot);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes, line);
    emit_decode_utf8(chunk, line);
}
