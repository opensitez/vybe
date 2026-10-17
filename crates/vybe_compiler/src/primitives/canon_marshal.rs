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
const BYTE_PUSH_BATCH: i32 = 254;

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
    emit_store_byte_array_impl(chunk, line, true, |chunk| {
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
    emit_store_byte_array_impl(chunk, line, false, push_array)
}

fn emit_store_byte_array_impl(
    chunk: &mut Chunk,
    line: u32,
    encoded_bytes: bool,
    push_array: impl FnOnce(&mut Chunk),
) -> (u16, u16) {
    let bytes = chunk.alloc_scratch(1);
    let len = chunk.alloc_scratch(1);
    let i = chunk.alloc_scratch(1);
    let encoded_byte = encoded_bytes.then(|| chunk.alloc_scratch(1));

    push_array(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, bytes, line);

    // len = bytes.length — the UTF-8 BYTE count, not the string's `.length`,
    // which counts UTF-16 code units.
    chunk.emit_op_u16(Op::LOCAL_GET, bytes, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len, line);

    let ptr = emit_alloc(chunk, line, len);

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
    chunk.emit_op(Op::I32_GE_U, line);
    chunk.emit_br_if(1, line);
    if let Some(byte) = encoded_byte {
        // Array.from(TextEncoder.encode(...)) contains only byte-valued
        // numbers. This canonical loop shape lets the VM copy the whole run
        // while remaining ordinary Wasm loads/stores in serialized output.
        chunk.emit_op_u16(Op::LOCAL_GET, bytes, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_i32_const(-1, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_SET, byte, line);
        chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_GET, byte, line);
    } else {
        chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_GET, bytes, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let to_i32b = chunk.add_import("wasm:js-number", "toI32");
        chunk.emit_call(to_i32b, 1, line);
    }
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
    let next = super::bits::raw_i64_scratch(chunk);
    let end = super::bits::raw_i64_scratch(chunk);

    crate::primitives::globals::emit_read(chunk, BUMP, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, base, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    let undef_test = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef_test, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_idx(Op::MEMORY_GROW, 0u32, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, base, line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if(line);
    chunk.emit_op(Op::UNREACHABLE, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_i32_const(65536, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base, line);
    chunk.emit_end(line);

    // The bump base is eight-byte aligned. Compute the widened end once for
    // both the next aligned allocation and the memory growth check.
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op(Op::I64_EXTEND_I32_U, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n_slot, line);
    chunk.emit_op(Op::I64_EXTEND_I32_U, line);
    chunk.emit_op(Op::I64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, end, line);
    chunk.emit_i64_const(15, line);
    chunk.emit_op(Op::I64_ADD, line);
    chunk.emit_i64_const(!7, line);
    chunk.emit_op(Op::I64_AND, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, next, line);
    chunk.emit_i64_const(u32::MAX as i64, line);
    chunk.emit_op(Op::I64_GT_U, line);
    chunk.emit_if(line);
    chunk.emit_op(Op::UNREACHABLE, line);
    chunk.emit_end(line);

    // Keep the allocation base fixed for buffers larger than one page.
    chunk.emit_op_u16(Op::LOCAL_GET, end, line);
    chunk.emit_op_idx(Op::MEMORY_SIZE, 0u32, line);
    chunk.emit_op(Op::I64_EXTEND_I32_U, line);
    chunk.emit_i64_const(16, line);
    chunk.emit_op(Op::I64_SHL, line);
    chunk.emit_op(Op::I64_GT_U, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, end, line);
    chunk.emit_i64_const(65535, line);
    chunk.emit_op(Op::I64_ADD, line);
    chunk.emit_i64_const(16, line);
    chunk.emit_op(Op::I64_SHR_U, line);
    chunk.emit_op_idx(Op::MEMORY_SIZE, 0u32, line);
    chunk.emit_op(Op::I64_EXTEND_I32_U, line);
    chunk.emit_op(Op::I64_SUB, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op_idx(Op::MEMORY_GROW, 0u32, line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if(line);
    chunk.emit_op(Op::UNREACHABLE, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ptr, line);

    // The checked, aligned next pointer also separates empty buffers.
    chunk.emit_op_u16(Op::LOCAL_GET, next, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
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
pub fn emit_append_bytes(
    chunk: &mut Chunk,
    line: u32,
    arr_slot: u16,
    ptr_slot: u16,
    len_slot: u16,
) {
    let i = chunk.alloc_scratch(1);
    let push = chunk.add_import("ecma:array", "push");
    let set = chunk.add_import("ecma:array", "set");
    let length = chunk.add_import("ecma:array", "length");
    let base = chunk.alloc_scratch(1);
    let cursor = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);
    let done = chunk.emit_block(line);
    emit_push_batch_loop(
        chunk,
        line,
        arr_slot,
        len_slot,
        i,
        cursor,
        push,
        BYTE_PUSH_BATCH,
    );
    emit_push_batch_loop(chunk, line, arr_slot, len_slot, i, cursor, push, 64);

    // Fewer than 64 remaining bytes use the fixed-arity setter.
    let (tail_copy, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_op(Op::I32_GE_U, line);
    chunk.emit_br_if(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_op(Op::I32_LOAD8_U, line);
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(tail_copy);
    chunk.emit_end(line);
    chunk.patch_block(done);
}

fn emit_push_batch_loop(
    chunk: &mut Chunk,
    line: u32,
    arr_slot: u16,
    len_slot: u16,
    i: u16,
    cursor: u16,
    push: u16,
    batch: i32,
) {
    let batch_done = chunk.emit_block(line);
    let (batch_copy, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_i32_const(batch, line);
    chunk.emit_op(Op::I32_LT_U, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    for offset in 0..batch {
        chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
        chunk.emit_op(Op::I32_LOAD8_U, line);
        if offset != 0 {
            // The existing optional memarg format mirrors Wasm's unsigned
            // effective-address offset. Avoid an i32 add for every byte.
            chunk.emit_leb_u32(0x80, line);
            chunk.emit_leb_u32(offset as u32, line);
        }
    }
    chunk.emit_call(push, (batch + 1) as u8, line);
    chunk.emit_op(Op::DROP, line); // push answers the new length

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(batch, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_i32_const(batch, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(batch_copy);
    chunk.emit_end(line);
    chunk.patch_block(batch_done);
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn byte_loops_compare_canonical_lengths_as_unsigned() {
        let mut store = Chunk::new("store-bytes");
        emit_store_byte_array(&mut store, 0, |chunk| emit_new_bytes(chunk, 0));
        let store_code = vybe_runtime::debug::disassemble(&store);
        assert!(store_code.contains("i32.ge_u"));
        assert!(!store_code.contains("i32.ge_s"));

        let mut append = Chunk::new("append-bytes");
        let array = append.alloc_scratch(1);
        let ptr = append.alloc_scratch(1);
        let len = append.alloc_scratch(1);
        emit_append_bytes(&mut append, 0, array, ptr, len);
        let append_code = vybe_runtime::debug::disassemble(&append);
        assert!(append_code.contains("i32.lt_u"));
        assert!(append_code.contains("i32.ge_u"));
        assert!(!append_code.contains("i32.gt_s"));
        assert!(!append_code.contains("i32.ge_s"));
    }

    #[test]
    fn byte_append_loop_emits_valid_wasm() {
        let mut chunk = Chunk::new("append-bytes-wasm");
        chunk.memory_min_pages.push(1);
        let array = chunk.alloc_scratch(1);
        let ptr = chunk.alloc_scratch(1);
        let len = chunk.alloc_scratch(1);
        emit_new_bytes(&mut chunk, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, array, 0);
        chunk.emit_i32_const(0, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, ptr, 0);
        chunk.emit_i32_const(0, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, len, 0);
        emit_append_bytes(&mut chunk, 0, array, ptr, len);
        chunk.emit_i32_const(0, 0);
        chunk.emit_op(Op::RETURN, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("canonical byte append must be valid standard Wasm");
    }

    #[test]
    fn byte_append_preserves_values_across_batch_boundaries() {
        use vybe_runtime::VM;
        use vybe_runtime::capabilities::Capabilities;
        use vybe_runtime::value::{ObjectKind, Value};

        for size in [
            0, 1, 63, 64, 65, 127, 253, 254, 255, 317, 318, 319, 508, 572, 262_144,
        ] {
            let mut chunk = Chunk::new("append-bytes-boundary");
            chunk.memory_min_pages.push(1);
            let array = chunk.alloc_scratch(1);
            let ptr = chunk.alloc_scratch(1);
            let len = chunk.alloc_scratch(1);
            emit_new_bytes(&mut chunk, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, array, 0);
            chunk.emit_i32_const(0, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, ptr, 0);
            chunk.emit_i32_const(size, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, len, 0);
            emit_append_bytes(&mut chunk, 0, array, ptr, len);
            chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
            chunk.emit_op(Op::RETURN, 0);

            let binary = matches!(size, 64 | 254 | 572 | 262_144)
                .then(|| vybe_platform_wasm::write_wasm(&[chunk.clone()]));
            if let Some(bytes) = &binary {
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(bytes)
                    .expect("canonical byte append must be valid standard Wasm");
            }

            let mut vm = VM::new();
            crate::primitives::platforms::register_platforms(&mut vm, &Capabilities::all());
            let expected: Vec<u8> = (0..size).map(|n| (n * 37 % 256) as u8).collect();
            vm.memory.grow(if size > 65_536 { 5 } else { 1 });
            vm.memory.write_bytes(0, &expected);
            let Value::Object(result) = vm.run(vec![chunk]).expect("byte append") else {
                panic!("byte append must return an array");
            };
            let guard = result.lock().unwrap();
            let ObjectKind::Array(values) = &guard.kind else {
                panic!("byte append must return a plain array");
            };
            let actual: Vec<u8> = values.iter().map(|value| value.as_i32() as u8).collect();
            assert_eq!(actual, expected, "size {size}");

            if let Some(bytes) = binary {
                let entry: Vec<_> = vybe_platform_wasm::read_wasm(&bytes)
                    .expect("read canonical byte append binary")
                    .into_iter()
                    .filter(|chunk| chunk.name == "_start")
                    .collect();
                assert_eq!(entry.len(), 1);
                let mut binary_vm = VM::new();
                crate::primitives::platforms::register_platforms(
                    &mut binary_vm,
                    &Capabilities::all(),
                );
                binary_vm.memory.grow(if size > 65_536 { 5 } else { 1 });
                binary_vm.memory.write_bytes(0, &expected);
                binary_vm.record_runtime_perf(true);
                let Value::Object(result) = binary_vm.run(entry).expect("binary byte append")
                else {
                    panic!("binary byte append must return an array");
                };
                let guard = result.lock().unwrap();
                let ObjectKind::Array(values) = &guard.kind else {
                    panic!("binary byte append must return a plain array");
                };
                let actual: Vec<u8> = values.iter().map(|value| value.as_i32() as u8).collect();
                assert_eq!(actual, expected, "binary size {size}");
                let expected_batches = u64::from(size >= 254) + u64::from(size % 254 >= 64);
                assert_eq!(
                    binary_vm
                        .take_runtime_perf()
                        .unwrap()
                        .super_linear_to_managed,
                    expected_batches,
                    "binary size {size}",
                );

                if size == 64 {
                    use std::sync::{Arc, Mutex};
                    use vybe_runtime::value::Object;

                    let mut guarded_vm = VM::new();
                    crate::primitives::platforms::register_platforms(
                        &mut guarded_vm,
                        &Capabilities::all(),
                    );
                    guarded_vm.register_host_fn(
                        "vybe:js-array",
                        "newWithLength",
                        Box::new(|_, _| {
                            let mut array = Object::new_array(Vec::new());
                            array.properties.insert("__proto__".into(), Value::Null);
                            array.properties.insert("marker".into(), Value::I32(1));
                            Value::Object(Arc::new(Mutex::new(array)))
                        }),
                    );
                    guarded_vm.memory.grow(1);
                    guarded_vm.memory.write_bytes(0, &expected);
                    guarded_vm.record_runtime_perf(true);
                    let Value::Object(result) = guarded_vm
                        .run(
                            vybe_platform_wasm::read_wasm(&bytes)
                                .expect("read guarded byte append binary")
                                .into_iter()
                                .filter(|chunk| chunk.name == "_start")
                                .collect(),
                        )
                        .expect("guarded byte append")
                    else {
                        panic!("guarded byte append must return an array");
                    };
                    let guard = result.lock().unwrap();
                    let ObjectKind::Array(values) = &guard.kind else {
                        panic!("guarded byte append must return a plain array");
                    };
                    let actual: Vec<u8> = values.iter().map(|value| value.as_i32() as u8).collect();
                    assert_eq!(actual, expected);
                    assert_eq!(
                        guarded_vm
                            .take_runtime_perf()
                            .unwrap()
                            .super_linear_to_managed,
                        0,
                    );
                }
            }
        }
    }

    #[test]
    fn allocator_checks_widened_next_pointer_before_growing_memory() {
        let mut chunk = Chunk::new("alloc-order");
        let n = chunk.alloc_scratch(1);
        emit_alloc(&mut chunk, 0, n);
        let code = vybe_runtime::debug::disassemble(&chunk);
        let guard = code.find("i64.gt_u").expect("widened pointer guard");
        let size = code.find("memory.size").expect("growth check");
        assert!(guard < size, "pointer overflow must trap before growth");
    }

    #[test]
    fn allocator_keeps_aligned_bump_positions_for_small_buffers() {
        use vybe_runtime::VM;
        use vybe_runtime::capabilities::Capabilities;

        for (size, next) in [
            (0, 65_544),
            (1, 65_552),
            (17, 65_568),
            (65_535, 131_080),
            (65_536, 131_080),
            (65_537, 131_088),
            (131_073, 196_624),
        ] {
            let mut chunk = Chunk::new("alloc-bump");
            chunk.memory_min_pages.push(1);
            let n = chunk.alloc_scratch(1);
            chunk.emit_i32_const(65_536, 0);
            crate::primitives::globals::emit_write(&mut chunk, BUMP, 0);
            chunk.emit_i32_const(size, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, n, 0);
            emit_alloc(&mut chunk, 0, n);
            crate::primitives::globals::emit_read(&mut chunk, BUMP, 0);
            chunk.emit_op(Op::RETURN, 0);

            let mut vm = VM::new();
            crate::primitives::platforms::register_platforms(&mut vm, &Capabilities::all());
            let result = vm.run(vec![chunk]).expect("canonical allocation");
            assert_eq!(result.as_i32(), next, "size {size}");
        }
    }
}
