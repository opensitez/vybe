use vybe_compiler::primitives::{collections, slices};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::emitter::adapter_util::{call_import, lget, lset, stash_args};

fn object_set_slot(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    key: &str,
    value: u16,
    line: u32,
) {
    lget(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(key, line);
    lget(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn object_set_string(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    key: &str,
    value: &str,
    line: u32,
) {
    lget(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_string_const(value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Python `slice(start[, stop[, step]])`.
///
/// The value is materialized as a tagged plain object, not a source prelude
/// class. That matches the XML adapter record shape and keeps `s.start` reads
/// on the same `__py_attr_read`/ECMA object path.
pub fn emit_slice_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 3 {
        return;
    }
    let base = stash_args(chunks, current, 3, line);
    let first = base;
    let second = base + 1;
    let third = base + 2;
    let actual_start = chunks[current].alloc_scratch(1);
    let actual_stop = chunks[current].alloc_scratch(1);
    let actual_step = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], second, line);
    vybe_compiler::primitives::globals::emit_read(&mut chunks[current], "__slice_unset", line);
    chunks[current].emit_op(Op::REF_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    lset(&mut chunks[current], actual_start, line);
    lget(&mut chunks[current], first, line);
    lset(&mut chunks[current], actual_stop, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    lset(&mut chunks[current], actual_step, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], first, line);
    lset(&mut chunks[current], actual_start, line);
    lget(&mut chunks[current], second, line);
    lset(&mut chunks[current], actual_stop, line);
    lget(&mut chunks[current], third, line);
    lset(&mut chunks[current], actual_step, line);
    chunks[current].emit_end(line);

    call_import(chunks, current, "ecma:object", "new", 0, line);
    let out = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], out, line);
    object_set_string(chunks, current, out, "__type", "slice", line);
    object_set_slot(chunks, current, out, "start", actual_start, line);
    object_set_slot(chunks, current, out, "stop", actual_stop, line);
    object_set_slot(chunks, current, out, "step", actual_step, line);
    lget(&mut chunks[current], out, line);
}

pub fn emit_getslice_obj(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 2 {
        return;
    }
    let base = stash_args(chunks, current, 2, line);
    let obj = base;
    let slice = base + 1;

    lget(&mut chunks[current], obj, line);
    for field in ["start", "stop", "step"] {
        lget(&mut chunks[current], slice, line);
        chunks[current].emit_string_const(field, line);
        call_import(chunks, current, "ecma:object", "get", 2, line);
    }
    slices::emit_stepped(chunks, current, line, slices::Options::new(true));
}
