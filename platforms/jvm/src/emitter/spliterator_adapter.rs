//! JVM `java.util.Spliterator` adapter.
//!
//! A spliterator over a list is a struct holding the backing ecma array, a
//! cursor, and its characteristics word. `trySplit` hands the FIRST half of
//! the remainder to a new spliterator and keeps the second (the JDK's
//! array-spliterator split), so parallel-shaped corpus code sees real
//! halving without any thread machinery.

use vybe_compiler::primitives::{callable, collections, instructions::host, ops};
use vybe_runtime::Chunk;
use vybe_runtime::Value;
use vybe_runtime::opcode::Op;

const ARR: &str = "__spl_arr";
const POS: &str = "__spl_pos";
const CHARS: &str = "__spl_chars";

/// `ORDERED | SIZED | SUBSIZED` — what `ArrayList.spliterator()` reports.
pub const LIST_CHARACTERISTICS: i64 = 0x0010 | 0x0040 | 0x4000;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn field_get(chunk: &mut Chunk, obj: u16, name: &str, line: u32) {
    get(chunk, obj, line);
    let k = chunk.add_constant(Value::String(name.into()));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

fn field_set_from_stack(chunk: &mut Chunk, obj: u16, name: &str, line: u32) {
    let value = chunk.alloc_scratch(1);
    set(chunk, value, line);
    get(chunk, obj, line);
    get(chunk, value, line);
    let k = chunk.add_constant(Value::String(name.into()));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}

/// `list.spliterator()` — `[array] -> [spliterator]`.
pub fn emit_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], arr, line);
    let obj = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    set(&mut chunks[current], obj, line);
    get(&mut chunks[current], arr, line);
    field_set_from_stack(&mut chunks[current], obj, ARR, line);
    chunks[current].emit_i32_const(0, line);
    field_set_from_stack(&mut chunks[current], obj, POS, line);
    chunks[current].emit_i64_const(LIST_CHARACTERISTICS, line);
    chunks[current].emit_op(Op::F64_CONVERT_I64_S, line);
    field_set_from_stack(&mut chunks[current], obj, CHARS, line);
    get(&mut chunks[current], obj, line);
}

/// Remaining element count, `[spl] -> [i32]` (helper, into `out`).
fn remaining_into(chunks: &mut [Chunk], current: usize, spl: u16, out: u16, line: u32) {
    field_get(&mut chunks[current], spl, ARR, line);
    collections::emit_len(chunks, current, line);
    field_get(&mut chunks[current], spl, POS, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    set(&mut chunks[current], out, line);
}

/// `spl.estimateSize()`.
pub fn emit_estimate_size(chunks: &mut [Chunk], current: usize, line: u32) {
    let spl = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], spl, line);
    let rem = chunks[current].alloc_scratch(1);
    remaining_into(chunks, current, spl, rem, line);
    get(&mut chunks[current], rem, line);
}

/// `spl.characteristics()`.
pub fn emit_characteristics(chunks: &mut [Chunk], current: usize, line: u32) {
    let spl = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], spl, line);
    field_get(&mut chunks[current], spl, CHARS, line);
}

/// `spl.hasCharacteristics(c)` — `(characteristics & c) == c`.
pub fn emit_has_characteristics(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = chunks[current].alloc_scratch(1);
    let spl = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], c, line);
    set(&mut chunks[current], spl, line);
    field_get(&mut chunks[current], spl, CHARS, line);
    chunks[current].emit_op(Op::I64_TRUNC_F64_S, line);
    get(&mut chunks[current], c, line);
    chunks[current].emit_op(Op::I64_TRUNC_F64_S, line);
    chunks[current].emit_op(Op::I64_AND, line);
    get(&mut chunks[current], c, line);
    chunks[current].emit_op(Op::I64_TRUNC_F64_S, line);
    chunks[current].emit_op(Op::I64_EQ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `spl.tryAdvance(consumer)` → whether an element was consumed.
pub fn emit_try_advance(chunks: &mut [Chunk], current: usize, line: u32) {
    let consumer = chunks[current].alloc_scratch(1);
    let spl = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], consumer, line);
    set(&mut chunks[current], spl, line);
    let rem = chunks[current].alloc_scratch(1);
    remaining_into(chunks, current, spl, rem, line);
    get(&mut chunks[current], rem, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], consumer, line);
    field_get(&mut chunks[current], spl, ARR, line);
    field_get(&mut chunks[current], spl, POS, line);
    collections::emit_get(chunks, current, line);
    callable::emit_direct_invoke(chunks, current, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    field_get(&mut chunks[current], spl, POS, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    field_set_from_stack(&mut chunks[current], spl, POS, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

/// `spl.forEachRemaining(consumer)` → null.
pub fn emit_for_each_remaining(chunks: &mut [Chunk], current: usize, line: u32) {
    let consumer = chunks[current].alloc_scratch(1);
    let spl = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], consumer, line);
    set(&mut chunks[current], spl, line);
    let rem = chunks[current].alloc_scratch(1);
    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    remaining_into(chunks, current, spl, rem, line);
    get(&mut chunks[current], rem, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], consumer, line);
    field_get(&mut chunks[current], spl, ARR, line);
    field_get(&mut chunks[current], spl, POS, line);
    collections::emit_get(chunks, current, line);
    callable::emit_direct_invoke(chunks, current, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    field_get(&mut chunks[current], spl, POS, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    field_set_from_stack(&mut chunks[current], spl, POS, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `spl.trySplit()` — first half of the remainder as a NEW spliterator, or
/// null when fewer than two elements remain.
pub fn emit_try_split(chunks: &mut [Chunk], current: usize, line: u32) {
    let spl = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], spl, line);
    let rem = chunks[current].alloc_scratch(1);
    remaining_into(chunks, current, spl, rem, line);
    get(&mut chunks[current], rem, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    let mid = chunks[current].alloc_scratch(1);
    field_get(&mut chunks[current], spl, POS, line);
    get(&mut chunks[current], rem, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_DIV_S, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], mid, line);
    field_get(&mut chunks[current], spl, ARR, line);
    field_get(&mut chunks[current], spl, POS, line);
    get(&mut chunks[current], mid, line);
    host::emit(&mut chunks[current], "ecma:array", "slice", 3, line);
    emit_new(chunks, current, line);
    let prefix = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], prefix, line);
    get(&mut chunks[current], mid, line);
    field_set_from_stack(&mut chunks[current], spl, POS, line);
    get(&mut chunks[current], prefix, line);
    chunks[current].emit_end(line);
}

/// `spl.getComparator()` — an unsorted spliterator throws (JDK contract).
pub fn emit_get_comparator(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    crate::emitter::exceptions::emit_jvm_exception_throw(
        chunks,
        current,
        "IllegalStateException",
        line,
    );
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

const STREAM_PARALLEL: &str = "__j_stream_parallel";

/// `StreamSupport.stream(spliterator, parallel)` — the spliterator's
/// remaining elements as a stream (an array in this runtime), carrying the
/// requested-parallel flag for `isParallel()`.
pub fn emit_stream_support_stream(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let parallel = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        set(&mut chunks[current], parallel, line);
    } else {
        chunks[current].emit_bool_const(false, line);
        set(&mut chunks[current], parallel, line);
    }
    let spl = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], spl, line);
    field_get(&mut chunks[current], spl, ARR, line);
    field_get(&mut chunks[current], spl, POS, line);
    host::emit(&mut chunks[current], "ecma:array", "slice", 2, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], out, line);
    chunks[current].emit_string_const(STREAM_PARALLEL, line);
    get(&mut chunks[current], parallel, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], out, line);
}

/// `stream.isParallel()` — the flag `StreamSupport.stream` recorded; a
/// stream built any other way answers false.
pub fn emit_stream_is_parallel(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const(STREAM_PARALLEL, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}
