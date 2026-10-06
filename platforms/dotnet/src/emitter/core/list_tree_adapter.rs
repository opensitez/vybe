//! WinForms tree and list items represented as ordinary DOM nodes.

use vybe_compiler::primitives::{
    gui::{DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT},
    ops,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::emitter::tree_register::TREE_NODE_MARKUP;

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

fn selection_style(chunk: &mut Chunk, node: u16, selected: bool, line: u32) {
    let details = first_child(chunk, node, line);
    let label = first_child(chunk, details, line);
    for (property, value) in [("background-color", "#0078d7"), ("color", "#ffffff")] {
        document(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, label, line);
        chunk.emit_string_const(property, line);
        if selected {
            chunk.emit_string_const(value, line);
            let set = chunk.add_import("web:cssom", "setStyleProperty");
            chunk.emit_call(set, 4, line);
        } else {
            let remove = chunk.add_import("web:cssom", "removeStyleProperty");
            chunk.emit_call(remove, 3, line);
        }
        chunk.emit_op(Op::DROP, line);
    }
}

fn first_child(chunk: &mut Chunk, parent: u16, line: u32) -> u16 {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, parent, line);
    dom(chunk, "firstChild", 2, line);
    let child = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, child, line);
    child
}

/// `TreeView.Nodes.Add(text)` -> a child `<li>`.
pub fn emit_tree_add(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let handler = tree_selection_callback(chunks, line);
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(3);
    let (tree, value, node) = (base, base + 1, base + 2);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, tree, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    dom(chunk, "nodeName", 2, line);
    chunk.emit_string_const("LI", line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    append(chunk, tree, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_else(line);
    create(chunk, "li", node, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const("role", line);
    chunk.emit_string_const("treeitem", line);
    dom(chunk, "setAttribute", 4, line);
    chunk.emit_op(Op::DROP, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const("style", line);
    chunk.emit_string_const("list-style:none;white-space:nowrap;padding:1px 4px", line);
    dom(chunk, "setAttribute", 4, line);
    chunk.emit_op(Op::DROP, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const(TREE_NODE_MARKUP, line);
    dom(chunk, "setInnerHtml", 3, line);
    chunk.emit_op(Op::DROP, line);
    let details = first_child(chunk, node, line);
    let label = first_child(chunk, details, line);
    text(chunk, label, value, line);
    append(chunk, tree, node, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_end(line);
    let added = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, added, line);
    let details = first_child(chunk, added, line);
    let label = first_child(chunk, details, line);
    super::tree_view_properties_adapter::apply_new_node(chunk, tree, label, line);
    vybe_compiler::primitives::gui::emit_add_event_listener(chunk, label, "click", handler, line);
    chunk.emit_op_u16(Op::LOCAL_GET, added, line);
}

fn tree_selection_callback(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut chunk =
        vybe_compiler::primitives::functions::create_function_chunk("__dotnet_tree_select", 1);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    chunk.emit_string_const("currentTarget", line);
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_call(get, 2, line);
    let summary = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, summary, line);
    document(&mut chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, summary, line);
    dom(&mut chunk, "parentNode", 2, line);
    let details = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, details, line);
    document(&mut chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, details, line);
    dom(&mut chunk, "parentNode", 2, line);
    let node = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, node, line);
    document(&mut chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const("[role=tree]", line);
    dom(&mut chunk, "closest", 3, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    emit_tree_selected_node_set(&mut chunk, line);
    chunk.emit_op(Op::RETURN, line);
    chunks.push(chunk);
    chunks.len() - 1
}

pub fn emit_tree_selected_node_get(chunk: &mut Chunk, line: u32) {
    let tree = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, tree, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tree, line);
    chunk.emit_string_const("[aria-selected=true]", line);
    dom(chunk, "querySelectorWithin", 3, line);
}

pub fn emit_tree_selected_node_set(chunk: &mut Chunk, line: u32) {
    let node = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, node, line);
    let tree = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, tree, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tree, line);
    emit_tree_selected_node_get(chunk, line);
    let previous = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, previous, line);
    chunk.emit_op_u16(Op::LOCAL_GET, previous, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_else(line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, previous, line);
    chunk.emit_string_const("aria-selected", line);
    dom(chunk, "removeAttribute", 3, line);
    chunk.emit_op(Op::DROP, line);
    selection_style(chunk, previous, false, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_else(line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tree, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    dom(chunk, "contains", 3, line);
    chunk.emit_if(line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const("aria-selected", line);
    chunk.emit_string_const("true", line);
    dom(chunk, "setAttribute", 4, line);
    chunk.emit_op(Op::DROP, line);
    selection_style(chunk, node, true, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_tree_clear(chunk: &mut Chunk, line: u32) {
    let tree = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, tree, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tree, line);
    chunk.emit_string_const("", line);
    dom(chunk, "setTextContent", 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_tree_remove(chunk: &mut Chunk, line: u32) {
    let node = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, node, line);
    let tree = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, tree, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tree, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    dom(chunk, "removeChild", 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `TreeNode.Text` is the label, not the combined text of all descendants.
pub fn emit_tree_node_text_get(chunk: &mut Chunk, line: u32) {
    let node = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, node, line);
    let details = first_child(chunk, node, line);
    let label = first_child(chunk, details, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, label, line);
    dom(chunk, "textContent", 2, line);
}

pub fn emit_tree_node_text_set(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    let node = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, node, line);
    let details = first_child(chunk, node, line);
    let label = first_child(chunk, details, line);
    text(chunk, label, value, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Lazily create the nested list so flat trees pay for only their visible rows.
pub fn emit_tree_node_children(chunk: &mut Chunk, line: u32) {
    let node = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, node, line);
    let details = first_child(chunk, node, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, details, line);
    dom(chunk, "children", 2, line);
    let children = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, children, line);
    chunk.emit_op_u16(Op::LOCAL_GET, children, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, children, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_else(line);
    let nested = chunk.alloc_scratch(1);
    create(chunk, "ul", nested, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nested, line);
    chunk.emit_string_const("role", line);
    chunk.emit_string_const("group", line);
    dom(chunk, "setAttribute", 4, line);
    chunk.emit_op(Op::DROP, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nested, line);
    chunk.emit_string_const("style", line);
    chunk.emit_string_const("list-style:none;margin:0;padding-left:16px", line);
    dom(chunk, "setAttribute", 4, line);
    chunk.emit_op(Op::DROP, line);
    append(chunk, details, nested, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nested, line);
    chunk.emit_end(line);
}

pub fn emit_tree_node_expand(chunk: &mut Chunk, open: bool, line: u32) {
    let node = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, node, line);
    let details = first_child(chunk, node, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, details, line);
    chunk.emit_string_const("open", line);
    if open {
        chunk.emit_string_const("", line);
        dom(chunk, "setAttribute", 4, line);
    } else {
        dom(chunk, "removeAttribute", 3, line);
    }
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_tree_node_is_expanded(chunk: &mut Chunk, line: u32) {
    let node = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, node, line);
    let details = first_child(chunk, node, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, details, line);
    chunk.emit_string_const("open", line);
    dom(chunk, "hasAttribute", 3, line);
}

/// `ListView.Items.Add(text)` -> a row in the list's `<tbody>`.
pub fn emit_list_add(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let handler = list_selection_callback(chunks, line);
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
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
    chunk.emit_string_const("style", line);
    chunk.emit_string_const("user-select:none", line);
    dom(chunk, "setAttribute", 4, line);
    chunk.emit_op(Op::DROP, line);
    text(chunk, cell, value, line);
    append(chunk, row, cell, line);
    append(chunk, body, row, line);
    vybe_compiler::primitives::gui::emit_add_event_listener(chunk, row, "click", handler, line);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
}

fn list_selection_callback(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut chunk =
        vybe_compiler::primitives::functions::create_function_chunk("__dotnet_list_select", 1);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    chunk.emit_string_const("currentTarget", line);
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_call(get, 2, line);
    let row = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, row, line);
    document(&mut chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
    dom(&mut chunk, "parentNode", 2, line);
    let body = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, body, line);
    document(&mut chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, body, line);
    chunk.emit_string_const("tr[aria-selected=true]", line);
    dom(&mut chunk, "querySelectorWithin", 3, line);
    let previous = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, previous, line);
    chunk.emit_op_u16(Op::LOCAL_GET, previous, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_else(line);
    document(&mut chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, previous, line);
    chunk.emit_string_const("aria-selected", line);
    dom(&mut chunk, "removeAttribute", 3, line);
    chunk.emit_op(Op::DROP, line);
    list_selection_style(&mut chunk, previous, false, line);
    chunk.emit_end(line);
    document(&mut chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
    chunk.emit_string_const("aria-selected", line);
    chunk.emit_string_const("true", line);
    dom(&mut chunk, "setAttribute", 4, line);
    chunk.emit_op(Op::DROP, line);
    list_selection_style(&mut chunk, row, true, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op(Op::RETURN, line);
    chunks.push(chunk);
    chunks.len() - 1
}

fn list_selection_style(chunk: &mut Chunk, row: u16, selected: bool, line: u32) {
    let cell = first_child(chunk, row, line);
    for (property, value) in [("background-color", "#0078d7"), ("color", "#ffffff")] {
        for target in [row, cell] {
            document(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_GET, target, line);
            chunk.emit_string_const(property, line);
            if selected {
                chunk.emit_string_const(value, line);
                let set = chunk.add_import("web:cssom", "setStyleProperty");
                chunk.emit_call(set, 4, line);
            } else {
                let remove = chunk.add_import("web:cssom", "removeStyleProperty");
                chunk.emit_call(remove, 3, line);
            }
            chunk.emit_op(Op::DROP, line);
        }
    }
}

/// The selected rows are the ListView's selected items, not its column headers.
pub fn emit_list_selected_items_get(chunk: &mut Chunk, line: u32) {
    let list = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, list, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, list, line);
    chunk.emit_string_const("tbody tr[aria-selected=true]", line);
    dom(chunk, "querySelectorAllWithin", 3, line);
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
