//! TreeView properties applied to its native nested-list markup.

use vybe_compiler::primitives::{loops, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

fn document(c: &mut Chunk, line: u32) {
    let f = c.add_import("web:html", "activeDocument");
    c.emit_call(f, 0, line);
}

fn property(c: &mut Chunk, tree: u16, name: &str, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, tree, line);
    c.emit_string_const(name, line);
    let f = c.add_import("ecma:object", "get");
    c.emit_call(f, 2, line);
}

fn store(c: &mut Chunk, tree: u16, name: &str, value: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, tree, line);
    c.emit_string_const(name, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let f = c.add_import("ecma:object", "set");
    c.emit_call(f, 3, line);
    c.emit_op(Op::DROP, line);
}

fn style(c: &mut Chunk, node: u16, name: &str, value: &str, line: u32) {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(name, line);
    c.emit_string_const(value, line);
    let f = c.add_import("web:cssom", "setStyleProperty");
    c.emit_call(f, 4, line);
    c.emit_op(Op::DROP, line);
}

fn apply_marker(c: &mut Chunk, summary: u16, visible: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, visible, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    style(c, summary, "display", "list-item", line);
    c.emit_else(line);
    style(c, summary, "display", "block", line);
    c.emit_end(line);
}

/// Stack: `[tree] -> [null]`.
pub fn emit_init(c: &mut Chunk, line: u32) {
    let tree = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, tree, line);
    for name in ["scrollable", "showplusminus"] {
        c.emit_bool_const(true, line);
        let value = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, value, line);
        store(c, tree, name, value, line);
    }
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: `[tree] -> [Boolean]`.
pub fn emit_get(c: &mut Chunk, name: &str, line: u32) {
    let tree = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, tree, line);
    property(c, tree, name, line);
}

/// Stack: `[tree, Boolean] -> [null]`.
pub fn emit_set(c: &mut Chunk, name: &str, line: u32) {
    let value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    let tree = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, tree, line);
    store(c, tree, name, value, line);
    if name == "scrollable" {
        c.emit_op_u16(Op::LOCAL_GET, value, line);
        ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        style(c, tree, "overflow", "auto", line);
        c.emit_else(line);
        style(c, tree, "overflow", "hidden", line);
        c.emit_end(line);
    } else {
        document(c, line);
        c.emit_op_u16(Op::LOCAL_GET, tree, line);
        c.emit_string_const("[role=treeitem] summary", line);
        let query = c.add_import("web:dom", "querySelectorAllWithin");
        c.emit_call(query, 3, line);
        let summaries = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, summaries, line);
        let index = c.alloc_scratch(1);
        c.emit_i32_const(0, line);
        c.emit_op_u16(Op::LOCAL_SET, index, line);
        let state = loops::emit_loop_start(std::slice::from_mut(c), 0, line);
        c.emit_op_u16(Op::LOCAL_GET, index, line);
        c.emit_op_u16(Op::LOCAL_GET, summaries, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        c.emit_op(Op::I32_LT_S, line);
        loops::emit_loop_cond_from_i32(std::slice::from_mut(c), 0, line);
        c.emit_op_u16(Op::LOCAL_GET, summaries, line);
        c.emit_op_u16(Op::LOCAL_GET, index, line);
        c.emit_op(Op::ARRAY_GET, line);
        let summary = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, summary, line);
        apply_marker(c, summary, value, line);
        c.emit_op_u16(Op::LOCAL_GET, index, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op_u16(Op::LOCAL_SET, index, line);
        loops::emit_loop_end(std::slice::from_mut(c), 0, state, line);
    }
    c.emit_ref_null(HT_EXTERN, line);
}

pub(super) fn apply_new_node(c: &mut Chunk, tree: u16, summary: u16, line: u32) {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, tree, line);
    c.emit_string_const("[role=tree]", line);
    let closest = c.add_import("web:dom", "closest");
    c.emit_call(closest, 3, line);
    let root = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, root, line);
    c.emit_op_u16(Op::LOCAL_GET, root, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    property(c, root, "showplusminus", line);
    let visible = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, visible, line);
    apply_marker(c, summary, visible, line);
    c.emit_end(line);
}
