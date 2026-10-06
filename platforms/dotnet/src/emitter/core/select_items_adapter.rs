//! WinForms list item collections backed by the select element's options.

use vybe_compiler::primitives::{loops, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const DOCUMENT: &str = "web:html";

fn document(chunk: &mut Chunk, line: u32) {
    let active = chunk.add_import(DOCUMENT, "activeDocument");
    chunk.emit_call(active, 0, line);
}

fn children(chunk: &mut Chunk, select: u16, line: u32) -> u16 {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, select, line);
    let f = chunk.add_import("web:dom", "children");
    chunk.emit_call(f, 2, line);
    let result = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    result
}

fn option_text(chunk: &mut Chunk, option: u16, line: u32) -> u16 {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, option, line);
    let f = chunk.add_import("web:dom", "textContent");
    chunk.emit_call(f, 2, line);
    let result = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    result
}

fn sorted(chunk: &mut Chunk, select: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, select, line);
    chunk.emit_string_const("sorted", line);
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_call(get, 2, line);
    ops::emit_dyn_to_bool(chunk, line);
}

fn move_before(chunk: &mut Chunk, select: u16, option: u16, prior: u16, line: u32) {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, select, line);
    chunk.emit_op_u16(Op::LOCAL_GET, option, line);
    chunk.emit_op_u16(Op::LOCAL_GET, prior, line);
    let insert = chunk.add_import("web:dom", "insertBefore");
    chunk.emit_call(insert, 4, line);
    chunk.emit_op(Op::DROP, line);
}

fn earlier_than(chunk: &mut Chunk, text: u16, prior: u16, line: u32) {
    let prior_text = option_text(chunk, prior, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    chunk.emit_op_u16(Op::LOCAL_GET, prior_text, line);
    let compare = chunk.add_import("ecma:string", "localeCompare");
    chunk.emit_call(compare, 2, line);
    chunk.emit_i32_const(0, line);
    ops::emit_dyn_lt(chunk, line);
}

fn place_option(chunk: &mut Chunk, select: u16, option: u16, limit: u16, line: u32) -> u16 {
    let text = option_text(chunk, option, line);
    let position = chunk.alloc_scratch(1);
    let result = chunk.alloc_scratch(1);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, position, line);
    chunk.emit_op_u16(Op::LOCAL_GET, limit, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    let nodes = children(chunk, select, line);
    let state = loops::emit_loop_start(std::slice::from_mut(chunk), 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, position, line);
    chunk.emit_op_u16(Op::LOCAL_GET, limit, line);
    chunk.emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(std::slice::from_mut(chunk), 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nodes, line);
    chunk.emit_op_u16(Op::LOCAL_GET, position, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let prior = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, prior, line);
    earlier_than(chunk, text, prior, line);
    chunk.emit_if(line);
    move_before(chunk, select, option, prior, line);
    chunk.emit_op_u16(Op::LOCAL_GET, position, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, limit, line);
    chunk.emit_op_u16(Op::LOCAL_SET, position, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, position, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, position, line);
    chunk.emit_end(line);
    loops::emit_loop_end(std::slice::from_mut(chunk), 0, state, line);
    result
}

/// Stack: `[select] -> [Boolean]`.
pub fn emit_get_sorted(chunk: &mut Chunk, line: u32) {
    let select = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, select, line);
    sorted(chunk, select, line);
    chunk.emit_if(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

/// Stack: `[select, Boolean] -> [null]`.
pub fn emit_set_sorted(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    let select = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, select, line);
    chunk.emit_op_u16(Op::LOCAL_GET, select, line);
    chunk.emit_string_const("sorted", line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let set = chunk.add_import("ecma:object", "set");
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    let snapshot = children(chunk, select, line);
    let index = chunk.alloc_scratch(1);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    let state = loops::emit_loop_start(std::slice::from_mut(chunk), 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, snapshot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(std::slice::from_mut(chunk), 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, snapshot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let option = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, option, line);
    let placed = place_option(chunk, select, option, index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, placed, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    loops::emit_loop_end(std::slice::from_mut(chunk), 0, state, line);
    chunk.emit_end(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Stack: `[select, item] -> [index]`.
pub fn emit_add(chunk: &mut Chunk, line: u32) {
    strings::emit_to_string(chunk, line);
    let text = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, text, line);
    let select = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, select, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, select, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    let add = chunk.add_import("web:html", "addItem");
    chunk.emit_call(add, 3, line);
    chunk.emit_op(Op::DROP, line);
    let nodes = children(chunk, select, line);
    let index = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, nodes, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    sorted(chunk, select, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, nodes, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let option = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, option, line);
    let placed = place_option(chunk, select, option, index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, placed, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
}

/// Stack: `[select] -> [number of options]`.
pub fn emit_count(chunk: &mut Chunk, line: u32) {
    let select = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, select, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, select, line);
    let children = chunk.add_import("web:dom", "children");
    chunk.emit_call(children, 2, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
}

/// Stack: `[select] -> [null]`.
pub fn emit_clear(chunk: &mut Chunk, line: u32) {
    let select = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, select, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, select, line);
    let clear = chunk.add_import("web:html", "clearItems");
    chunk.emit_call(clear, 2, line);
}
