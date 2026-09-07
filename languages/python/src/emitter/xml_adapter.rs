//! Python `xml.etree.ElementTree` runtime adapters over the normalized element
//! object shape the walker emits.

use vybe_compiler::primitives::{collections, loops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

use super::adapter_util::{call_import, stash_exact};

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn object_get(chunks: &mut [Chunk], current: usize, object: u16, field: &str, line: u32) {
    lget(&mut chunks[current], object, line);
    chunks[current].emit_string_const(field, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
}

fn object_set_slot(
    chunks: &mut [Chunk],
    current: usize,
    object: u16,
    field: &str,
    value: u16,
    line: u32,
) {
    lget(&mut chunks[current], object, line);
    chunks[current].emit_string_const(field, line);
    lget(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn object_set_string(
    chunks: &mut [Chunk],
    current: usize,
    object: u16,
    field: &str,
    value: &str,
    line: u32,
) {
    lget(&mut chunks[current], object, line);
    chunks[current].emit_string_const(field, line);
    chunks[current].emit_string_const(value, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn object_set_null(chunks: &mut [Chunk], current: usize, object: u16, field: &str, line: u32) {
    lget(&mut chunks[current], object, line);
    chunks[current].emit_string_const(field, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    call_import(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn slot_is_missing(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "wasm:js-undefined", "test", 1, line);
    chunks[current].emit_op(Op::I32_OR, line);
}

fn tag_matches(chunks: &mut [Chunk], current: usize, elem: u16, tag: u16, line: u32) {
    lget(&mut chunks[current], tag, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    lget(&mut chunks[current], tag, line);
    call_import(chunks, current, "wasm:js-undefined", "test", 1, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    {
        chunks[current].emit_i32_const(1, line);
    }
    chunks[current].emit_else(line);
    {
        lget(&mut chunks[current], tag, line);
        chunks[current].emit_string_const("*", line);
        call_import(chunks, current, "wasm:js-string", "compare", 2, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if_value(line);
        {
            chunks[current].emit_i32_const(1, line);
        }
        chunks[current].emit_else(line);
        {
            object_get(chunks, current, elem, "tag", line);
            lget(&mut chunks[current], tag, line);
            call_import(chunks, current, "wasm:js-string", "compare", 2, line);
            chunks[current].emit_op(Op::I32_EQZ, line);
        }
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
}

fn push_if_matches(
    chunks: &mut [Chunk],
    current: usize,
    out: u16,
    elem: u16,
    tag: u16,
    line: u32,
) {
    tag_matches(chunks, current, elem, tag, line);
    chunks[current].emit_if(line);
    {
        lget(&mut chunks[current], out, line);
        lget(&mut chunks[current], elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_end(line);
}

fn append_direct_children(
    chunks: &mut [Chunk],
    current: usize,
    out: u16,
    elem: u16,
    tag: u16,
    line: u32,
) {
    let children = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let child = chunks[current].alloc_scratch(1);

    object_get(chunks, current, elem, "_children", line);
    lset(&mut chunks[current], children, line);
    let state = loops::emit_for_in_start(chunks, current, children, idx, line);
    lset(&mut chunks[current], child, line);

    push_if_matches(chunks, current, out, child, tag, line);

    loops::emit_for_in_end(chunks, current, idx, state, line);
}

fn emit_new_element_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    tag: u16,
    attrib: u16,
    line: u32,
) -> u16 {
    let elem = chunks[current].alloc_scratch(1);
    let attrib_value = chunks[current].alloc_scratch(1);
    let empty_children = chunks[current].alloc_scratch(1);

    call_import(chunks, current, "ecma:object", "new", 0, line);
    lset(&mut chunks[current], elem, line);
    object_set_string(chunks, current, elem, "__type", "xml_element", line);
    object_set_slot(chunks, current, elem, "tag", tag, line);

    lget(&mut chunks[current], attrib, line);
    lset(&mut chunks[current], attrib_value, line);
    slot_is_missing(chunks, current, attrib_value, line);
    chunks[current].emit_if(line);
    call_import(chunks, current, "ecma:object", "new", 0, line);
    lset(&mut chunks[current], attrib_value, line);
    chunks[current].emit_end(line);
    object_set_slot(chunks, current, elem, "attrib", attrib_value, line);

    object_set_null(chunks, current, elem, "text", line);
    object_set_string(chunks, current, elem, "tail", "", line);
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], empty_children, line);
    object_set_slot(chunks, current, elem, "_children", empty_children, line);
    elem
}

/// `Element(tag[, attrib])`.
/// Stack: `[tag, attrib?] -> [element]`.
pub fn emit_element(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let elem = emit_new_element_from_slots(chunks, current, base, base + 1, line);
    lget(&mut chunks[current], elem, line);
}

/// `SubElement(parent, tag[, attrib])`.
/// Stack: `[parent, tag, attrib?] -> [element]`.
pub fn emit_subelement(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 3, line);
    let parent = base;
    let tag = base + 1;
    let attrib = base + 2;
    let elem = emit_new_element_from_slots(chunks, current, tag, attrib, line);
    let children = chunks[current].alloc_scratch(1);

    object_get(chunks, current, parent, "_children", line);
    lset(&mut chunks[current], children, line);
    lget(&mut chunks[current], children, line);
    lget(&mut chunks[current], elem, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], elem, line);
}

/// `Element.find(tag)` for the normalized object shape.
/// Stack: `[elem, tag] -> [element-or-null]`.
pub fn emit_find(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let elem = base;
    let tag = base + 1;
    let children = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let child = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);

    chunks[current].emit_ref_null(HT_EXTERN, line);
    lset(&mut chunks[current], result, line);
    object_get(chunks, current, elem, "_children", line);
    lset(&mut chunks[current], children, line);

    let state = loops::emit_for_in_start(chunks, current, children, idx, line);
    lset(&mut chunks[current], child, line);

    lget(&mut chunks[current], result, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    tag_matches(chunks, current, child, tag, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], child, line);
    lset(&mut chunks[current], result, line);
    chunks[current].emit_end(line);

    loops::emit_for_in_end(chunks, current, idx, state, line);
    lget(&mut chunks[current], result, line);
}

/// `Element.findall(tag)`.
/// Stack: `[elem, tag] -> [array]`.
pub fn emit_findall(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let elem = base;
    let tag = base + 1;
    let out = chunks[current].alloc_scratch(1);

    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], out, line);
    append_direct_children(chunks, current, out, elem, tag, line);
    lget(&mut chunks[current], out, line);
}

/// `Element.iter([tag])`.
/// Stack: `[elem, tag?] -> [array]`.
pub fn emit_iter(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let elem = base;
    let tag = base + 1;
    let out = chunks[current].alloc_scratch(1);
    let work = chunks[current].alloc_scratch(1);
    let scan = chunks[current].alloc_scratch(1);
    let node = chunks[current].alloc_scratch(1);
    let children = chunks[current].alloc_scratch(1);
    let child_idx = chunks[current].alloc_scratch(1);
    let child = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);

    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], out, line);
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], work, line);

    lget(&mut chunks[current], work, line);
    lget(&mut chunks[current], elem, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], scan, line);

    let state = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], work, line);
    collections::emit_len(chunks, current, line);
    lset(&mut chunks[current], len, line);
    lget(&mut chunks[current], scan, line);
    lget(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], work, line);
    lget(&mut chunks[current], scan, line);
    collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], node, line);
    push_if_matches(chunks, current, out, node, tag, line);

    object_get(chunks, current, node, "_children", line);
    lset(&mut chunks[current], children, line);
    let child_loop = loops::emit_for_in_start(chunks, current, children, child_idx, line);
    lset(&mut chunks[current], child, line);
    lget(&mut chunks[current], work, line);
    lget(&mut chunks[current], child, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, child_idx, child_loop, line);

    lget(&mut chunks[current], scan, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], scan, line);
    loops::emit_loop_end(chunks, current, state, line);

    lget(&mut chunks[current], out, line);
}

/// `Element.get(key[, default])`.
/// Stack: `[elem, key, default?] -> [value-or-default]`.
pub fn emit_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 3, line);
    let elem = base;
    let key = base + 1;
    let default = base + 2;
    let value = chunks[current].alloc_scratch(1);

    object_get(chunks, current, elem, "attrib", line);
    lget(&mut chunks[current], key, line);
    call_import(chunks, current, "ecma:object", "get", 2, line);
    lset(&mut chunks[current], value, line);

    slot_is_missing(chunks, current, value, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], default, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}
