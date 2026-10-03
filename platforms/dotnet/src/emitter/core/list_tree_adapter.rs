//! WinForms tree and list items represented as ordinary DOM nodes.

use vybe_compiler::primitives::gui::{DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const DOM: &str = "web:dom";

fn document(chunk: &mut Chunk, line: u32) {
    let import = chunk.add_import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(import, 0, line);
}

fn dom(chunk: &mut Chunk, name: &str, argc: u8, line: u32) {
    let import = chunk.add_import(DOM, name);
    chunk.emit_call(import, argc, line);
}

fn create(chunk: &mut Chunk, tag: &str, slot: u16, line: u32) {
    document(chunk, line);
    chunk.emit_string_const(tag, line);
    chunk.emit_string_const("", line);
    dom(chunk, "createElement", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn append(chunk: &mut Chunk, parent: u16, child: u16, line: u32) {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, parent, line);
    chunk.emit_op_u16(Op::LOCAL_GET, child, line);
    dom(chunk, "appendChild", 3, line);
    chunk.emit_op(Op::DROP, line);
}

fn text(chunk: &mut Chunk, node: u16, value: u16, line: u32) {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    dom(chunk, "setTextContent", 3, line);
    chunk.emit_op(Op::DROP, line);
}

/// `TreeView.Nodes.Add(text)` -> a child `<li>`.
pub fn emit_tree_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(3);
    let (tree, value, node) = (base, base + 1, base + 2);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, tree, line);
    create(chunk, "li", node, line);
    text(chunk, node, value, line);
    append(chunk, tree, node, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
}

/// `ListView.Items.Add(text)` -> a row in the list's `<tbody>`.
pub fn emit_list_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(5);
    let (list, value, body, row, cell) = (base, base + 1, base + 2, base + 3, base + 4);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, list, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, list, line);
    dom(chunk, "lastChild", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, body, line);
    create(chunk, "tr", row, line);
    create(chunk, "td", cell, line);
    text(chunk, cell, value, line);
    append(chunk, row, cell, line);
    append(chunk, body, row, line);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
}

/// Count the DOM items held by a TreeView, ListView section, or TabControl.
pub fn emit_count(chunk: &mut Chunk, section: &str, line: u32) {
    let root = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, root, line);
    let mut current = root;
    for edge in match section {
        "columns" => &["firstChild", "firstChild"][..],
        "items" => &["lastChild"][..],
        _ => &[][..],
    } {
        let child = chunk.alloc_scratch(1);
        document(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, current, line);
        dom(chunk, edge, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, child, line);
        current = child;
    }
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, current, line);
    dom(chunk, "children", 2, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
}
