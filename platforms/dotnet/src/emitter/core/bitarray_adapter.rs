//! `System.Collections.BitArray` construction and word packing.
//!
//! Only the parts that are specific to .NET's API live here. The rest of the
//! type is bound as leaves in `component_classes_collections_specialized.rs`:
//! the indexer, `Length`, `Count`, `Clone` and `CopyTo` are the underlying
//! array's own members, `SetAll` is `collections.fill_all`, and
//! `And`/`Or`/`Xor`/`Not` are `bits.array_*`.
//!
//! A `BitArray` IS a boolean array. Everything below converts between that and
//! the packed-word forms .NET also accepts.

use vybe_compiler::primitives::instructions::host;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Width in bits of one element of a packed source array.
///
/// `BitArray(byte[])` yields 8 bits per element and `BitArray(int[])` 32, and
/// the declared element type is gone by the time this runs. The largest value
/// present decides: any element above 255 cannot be a byte.
fn emit_word_width(chunks: &mut [Chunk], current: usize, src: u16, out: u16, line: u32) {
    let idx = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];

    chunk.emit_i32_const(8, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);

    chunk.emit_op_u16(Op::LOCAL_GET, src, line);
    host::emit(chunk, "ecma:array", "length", 1, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx, line);

    let done = chunk.emit_block(line);
    let (again, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, src, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx, line);
    host::emit(chunk, "ecma:array", "get", 2, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_i32_const(255, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, idx, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(again);
    chunk.emit_end(line);
    chunk.patch_block(done);
}

/// Expand a packed numeric array into one boolean per bit, least significant
/// bit of each word first. Stack: `[]` → `[]`, result left in `out`.
fn emit_unpack(chunks: &mut [Chunk], current: usize, src: u16, out: u16, line: u32) {
    let width = chunks[current].alloc_scratch(1);
    let word = chunks[current].alloc_scratch(1);
    let bit = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);

    emit_word_width(chunks, current, src, width, line);
    let chunk = &mut chunks[current];

    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);

    chunk.emit_op_u16(Op::LOCAL_GET, src, line);
    host::emit(chunk, "ecma:array", "length", 1, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, word, line);

    let outer_done = chunk.emit_block(line);
    let (outer, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, bit, line);

    let inner_done = chunk.emit_block(line);
    let (inner, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, bit, line);
    chunk.emit_op_u16(Op::LOCAL_GET, width, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, src, line);
    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
    host::emit(chunk, "ecma:array", "get", 2, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bit, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
    host::emit(chunk, "ecma:array", "push", 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, bit, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, bit, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(inner);
    chunk.emit_end(line);
    chunk.patch_block(inner_done);

    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, word, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(outer);
    chunk.emit_end(line);
    chunk.patch_block(outer_done);
}

/// `new BitArray(…)`.
///
/// One entry for all three documented shapes: a bit COUNT, a boolean array to
/// copy, or a packed `byte[]`/`int[]` to expand. A `ClassType` holds exactly
/// one constructor, and the call site passes the real argument count, so the
/// arity is resolved here rather than by declaring several.
///
/// The one-argument form is ambiguous until run time — `BitArray(4)` is four
/// false bits, `BitArray(@($true,$false))` is those two — so the operand is
/// tested rather than trusted to a declared type. A boolean array is COPIED:
/// `new BitArray(bits)` does not alias its source.
///
/// Stack: `[count]` / `[bits]` / `[words]` / `[count, value]` → `[array]`.
pub fn emit_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 2 {
        let chunk = &mut chunks[current];
        let value = chunk.alloc_scratch(1);
        let count = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
        chunk.emit_op_u16(Op::LOCAL_SET, count, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        chunk.emit_op_u16(Op::LOCAL_GET, count, line);
        vybe_compiler::primitives::collections::emit_repeat_value(chunks, current, line);
        return;
    }

    let arg = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let unpacked = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arg, line);

    // Each branch writes `out` and the result is pushed once at the end. The
    // repeat and the unpack open blocks of their own and branch out by depth,
    // so neither can sit inside a value-producing `if`.
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    vybe_compiler::primitives::collections::emit_repeat_value(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    emit_unpack(chunks, current, arg, unpacked, line);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, arg, line);
    host::emit(chunk, "ecma:array", "isArray", 1, line);
    host::emit(chunk, "wasm:js-boolean", "cast", 1, line);
    chunk.emit_if(line);

    // A boolean array is the bits themselves; a numeric one is packed words.
    chunk.emit_op_u16(Op::LOCAL_GET, arg, line);
    chunk.emit_i32_const(0, line);
    host::emit(chunk, "ecma:array", "get", 2, line);
    host::emit(chunk, "wasm:js-boolean", "test", 1, line);
    chunk.emit_if(line);
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arg, line);
    host::emit(chunk, "ecma:array", "concat", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, unpacked, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_end(line);

    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `CopyTo(target, index)` — pack the bits back into `target`'s words.
///
/// The element width is taken from `target`'s length against the bit count:
/// 32 bits packed into one element means an `int[]`, 8 means a `byte[]`.
///
/// Stack: `[bits, target, index]` → `[target]`.
pub fn emit_copy_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let index = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let bits = chunks[current].alloc_scratch(1);
    let width = chunks[current].alloc_scratch(1);
    let bitlen = chunks[current].alloc_scratch(1);
    let word = chunks[current].alloc_scratch(1);
    let bit = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    let tlen = chunks[current].alloc_scratch(1);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target, line);
    chunk.emit_op_u16(Op::LOCAL_SET, bits, line);

    chunk.emit_op_u16(Op::LOCAL_GET, bits, line);
    host::emit(chunk, "ecma:array", "length", 1, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, bitlen, line);

    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    host::emit(chunk, "ecma:array", "length", 1, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, tlen, line);

    chunk.emit_i32_const(8, line);
    chunk.emit_op_u16(Op::LOCAL_SET, width, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bitlen, line);
    chunk.emit_i32_const(8, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op_u16(Op::LOCAL_SET, width, line);
    chunk.emit_end(line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, word, line);

    let outer_done = chunk.emit_block(line);
    let (outer, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tlen, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, bit, line);

    let inner_done = chunk.emit_block(line);
    let (inner, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, bit, line);
    chunk.emit_op_u16(Op::LOCAL_GET, width, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    // `acc |= 1 << bit` when the bit at `word * width + bit` is set and in range.
    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
    chunk.emit_op_u16(Op::LOCAL_GET, width, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bit, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bitlen, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, bits, line);
    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
    chunk.emit_op_u16(Op::LOCAL_GET, width, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bit, line);
    chunk.emit_op(Op::I32_ADD, line);
    host::emit(chunk, "ecma:array", "get", 2, line);
    host::emit(chunk, "wasm:js-boolean", "cast", 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bit, line);
    chunk.emit_op(Op::I32_SHL, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);
    chunk.emit_end(line);

    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, bit, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, bit, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(inner);
    chunk.emit_end(line);
    chunk.patch_block(inner_done);

    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    host::emit(chunk, "ecma:array", "set", 3, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, word, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(outer);
    chunk.emit_end(line);
    chunk.patch_block(outer_done);

    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
}

/// `$ba.Length = n` — grow with `false` or truncate in place.
///
/// Stack: `[bits, n]` → `[bits]`.
pub fn emit_resize(chunks: &mut [Chunk], current: usize, line: u32) {
    let want = chunks[current].alloc_scratch(1);
    let bits = chunks[current].alloc_scratch(1);
    let have = chunks[current].alloc_scratch(1);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, want, line);
    chunk.emit_op_u16(Op::LOCAL_SET, bits, line);

    chunk.emit_op_u16(Op::LOCAL_GET, bits, line);
    host::emit(chunk, "ecma:array", "length", 1, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, have, line);

    let done = chunk.emit_block(line);
    let (again, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, have, line);
    chunk.emit_op_u16(Op::LOCAL_GET, want, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, bits, line);
    chunk.emit_bool_const(false, line);
    host::emit(chunk, "ecma:array", "push", 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, have, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, have, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(again);
    chunk.emit_end(line);
    chunk.patch_block(done);

    // Truncation is a splice from the requested length to the end.
    chunk.emit_op_u16(Op::LOCAL_GET, bits, line);
    chunk.emit_op_u16(Op::LOCAL_GET, want, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, have, line);
    host::emit(chunk, "ecma:array", "splice", 3, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, bits, line);
}

/// `IsSynchronized` — a `BitArray` is never synchronized.
///
/// Stack: `[bits]` → `[false]`.
pub fn emit_is_synchronized(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_bool_const(false, line);
}
