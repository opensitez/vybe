//! Getting guest data INTO linear memory, for the canonical ABI.
//!
//! The Component Model moves values through **linear memory**: `canon
//! stream.write` is `(handle, ptr, n)` and reads `n` elements starting at
//! `ptr`. Every other emitter in this compiler works on the GC/Value model —
//! a string is a `Value::String`, an array is a heap object — and until now
//! linear memory was touched in exactly one place, `channels.rs`, for futex
//! WORDS. There was no way to put a string's bytes anywhere a conforming
//! component could read them.
//!
//! That absence is why `canon stream.write` still accepted a Value item: not a
//! design decision, just the only shape reachable without this file. It is
//! also why `canon_value::store` refuses `string` and `list` — the spec calls
//! that allocation `realloc`, and this module is its analogue.
//!
//! Two helpers, both `__stdlib_*` runtime helpers linked once per program:
//!
//! - `__vybe_canon_alloc(n)` → a bump pointer into the page the futex
//!   allocator already grows. Generalises `__vybe_futex_alloc16`, which is
//!   arity-0, fixed at 16 bytes, and zeroes a channel-specific word.
//! - `__vybe_canon_store_utf8(s)` → encodes a string to UTF-8, allocates,
//!   copies, and answers the `(ptr, len)` pair packed into one i64 — the same
//!   packing `canon stream.new` uses for its two handles, for the same reason:
//!   a helper returns one value.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::primitives::collections;

/// `__vybe_canon_alloc(n) -> i32` — bump `n` bytes and answer the base.
///
/// Shares `__vybe_chan_futex_next` with the channel allocator on purpose:
/// two bump pointers into one page would hand out the same address twice.
/// Sizes are rounded up to 8 so a later `i64`/`f64` store lands aligned —
/// the canonical ABI asserts alignment, and an allocator that ignores it just
/// moves the trap somewhere less obvious.
pub fn build_canon_alloc(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_canon_alloc");
    c.arity = 1;
    c.local_count = 3; // n, base, size
    let (n, base, size) = (0u16, 1u16, 2u16);
    let line = 0u32;

    // size = align_to(n, 8), and never zero — a zero-length allocation must
    // still yield a distinct address, or two empty buffers alias.
    c.emit_op_u16(Op::LOCAL_GET, n, line);
    c.emit_i32_const(7, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_i32_const(!7, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op_u16(Op::LOCAL_TEE, size, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_if(line);
    c.emit_i32_const(8, line);
    c.emit_op_u16(Op::LOCAL_SET, size, line);
    c.emit_end(line);

    // Grow a page the first time, exactly as the futex allocator does.
    crate::primitives::globals::emit_read(&mut c, "__vybe_chan_futex_next", line);
    c.emit_op_u16(Op::LOCAL_TEE, base, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op_u16(Op::LOCAL_GET, base, line);
    collections::emit_import_call_into(imports, &mut c, "wasm:js-undefined", "test", 1, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_if(line);
    c.emit_i32_const(1, line);
    c.emit_op_u16(Op::MEMORY_GROW, 0, line);
    c.emit_i32_const(65536, line);
    c.emit_op(Op::I32_MUL, line);
    c.emit_op_u16(Op::LOCAL_SET, base, line);
    c.emit_end(line);

    // Grow again when this request would run off the end of the page. Without
    // this the allocator silently hands out addresses past the memory it owns,
    // and the resulting corruption surfaces nowhere near here.
    c.emit_op_u16(Op::LOCAL_GET, base, line);
    c.emit_op_u16(Op::LOCAL_GET, size, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::MEMORY_SIZE, 0, line);
    c.emit_i32_const(65536, line);
    c.emit_op(Op::I32_MUL, line);
    c.emit_op(Op::I32_GT_U, line);
    c.emit_if(line);
    c.emit_i32_const(1, line);
    c.emit_op_u16(Op::MEMORY_GROW, 0, line);
    c.emit_i32_const(65536, line);
    c.emit_op(Op::I32_MUL, line);
    c.emit_op_u16(Op::LOCAL_SET, base, line);
    c.emit_end(line);

    // next = base + size
    c.emit_op_u16(Op::LOCAL_GET, base, line);
    c.emit_op_u16(Op::LOCAL_GET, size, line);
    c.emit_op(Op::I32_ADD, line);
    crate::primitives::globals::emit_write(&mut c, "__vybe_chan_futex_next", line);

    c.emit_op_u16(Op::LOCAL_GET, base, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__vybe_canon_store_utf8(s) -> i64` — write a string's UTF-8 bytes into
/// linear memory; answer `ptr | (len << 32)`.
///
/// The byte count is the UTF-8 length, NOT the string's `.length`, which is
/// UTF-16 code units — they differ for anything outside the BMP's ASCII range,
/// and using the wrong one truncates or over-reads. `web:encoding`'s
/// TextEncoder is the same encoder the `utf8.encode` adapters already use, so
/// there is one definition of "the bytes of this string" in the tree.
///
/// Packed into an i64 for the same reason `canon stream.new` packs its handles:
/// a call answers one value, and (ptr, len) is inseparable — handing back a
/// pointer whose length the caller has to recompute is how the two drift.
pub fn build_canon_store_utf8(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_canon_store_utf8");
    c.arity = 1;
    c.local_count = 5; // s, bytes, len, ptr, i
    let (s, bytes, len, ptr, i) = (0u16, 1u16, 2u16, 3u16, 4u16);
    let line = 0u32;

    // bytes = Array.from(TextEncoder().encode(s))
    //
    // ⚠ The `Array.from` is LOAD-BEARING, not tidiness. `web:encoding.encode`
    // answers a `Uint8Array` — `ObjectKind::TypedArray` — and NOTHING in reach
    // indexes one: `Op::ARRAY_GET` matches only `ObjectKind::Array`, and
    // neither `ecma:array.at` nor plain indexing handles the typed case.
    // Without this conversion every byte read as 0, so the write emitted the
    // right COUNT of NUL bytes and stdout looked simply EMPTY — python and php
    // printed nothing at all while exiting 0.
    //
    // `ecma:array.from` is the same conversion `strings::emit_scalar_chars`
    // uses, so there is one answer in the tree for "make this indexable".
    collections::emit_import_call_into(imports, &mut c, "web:encoding", "encoderNew", 0, line);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    collections::emit_import_call_into(imports, &mut c, "web:encoding", "encode", 2, line);
    collections::emit_import_call_into(imports, &mut c, "ecma:array", "from", 1, line);
    c.emit_op_u16(Op::LOCAL_SET, bytes, line);

    // len = bytes.length — the UTF-8 BYTE count the encoder produced, not the
    // string's `.length`, which counts UTF-16 code units.
    c.emit_op_u16(Op::LOCAL_GET, bytes, line);
    collections::emit_import_call_into(imports, &mut c, "ecma:array", "length", 1, line);
    collections::emit_import_call_into(imports, &mut c, "wasm:js-number", "toI32", 1, line);
    c.emit_op_u16(Op::LOCAL_SET, len, line);

    // ptr = __vybe_canon_alloc(len)
    c.emit_op_u16(Op::LOCAL_GET, len, line);
    crate::primitives::bundle::emit_call_push_func(&mut c, "__vybe_canon_alloc", line);
    c.emit_op_u16(Op::LOCAL_SET, ptr, line);

    // for i in 0..len { i32.store8(ptr + i, bytes[i]) }
    //
    // `block { loop { br_if 1 (done); …; br 0 } }` — depth 1 leaves the block,
    // depth 0 re-enters the loop. `I32_STORE8` carries NO memarg: the VM's is
    // marker-tagged and optional, absent meaning natural align, offset 0,
    // memory 0 — which is what a byte store wants.
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);
    let _done_block = c.emit_block(line);
    let (_copy_loop, _) = c.emit_loop_s(line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op_u16(Op::LOCAL_GET, len, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);
    c.emit_op_u16(Op::LOCAL_GET, ptr, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_GET, bytes, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op(Op::ARRAY_GET, line);
    collections::emit_import_call_into(imports, &mut c, "wasm:js-number", "toI32", 1, line);
    c.emit_op(Op::I32_STORE8, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.emit_end(line);

    // ptr | (len << 32)
    c.emit_op_u16(Op::LOCAL_GET, ptr, line);
    c.emit_op(Op::I64_EXTEND_I32_U, line);
    c.emit_op_u16(Op::LOCAL_GET, len, line);
    c.emit_op(Op::I64_EXTEND_I32_U, line);
    c.emit_i64_const(32, line);
    c.emit_op(Op::I64_SHL, line);
    c.emit_op(Op::I64_OR, line);
    c.emit_op(Op::RETURN, line);
    c
}
