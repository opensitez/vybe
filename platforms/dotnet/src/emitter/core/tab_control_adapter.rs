//! WinForms TabControl pages and selection on ordinary DOM elements.

use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::{globals, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

pub fn emit_page_text_set(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    let page = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, page, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, page, line);
    chunk.emit_string_const("data-text", line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let attribute = chunk.add_import("web:dom", "setAttribute");
    chunk.emit_call(attribute, 4, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, page, line);
    chunk.emit_string_const("__dotnet_tab_button", line);
    let has = chunk.add_import("ecma:object", "hasOwn");
    chunk.emit_call(has, 2, line);
    chunk.emit_if(line);
    field_get(chunk, page, "__dotnet_tab_button", line);
    let button = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, button, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, button, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let set_text = chunk.add_import("web:dom", "setTextContent");
    chunk.emit_call(set_text, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_ref_null(HT_EXTERN, line);
}

pub fn emit_page_text_get(chunk: &mut Chunk, line: u32) {
    let page = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, page, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, page, line);
    chunk.emit_string_const("data-text", line);
    let get = chunk.add_import("web:dom", "getAttribute");
    chunk.emit_call(get, 3, line);
}

fn document(chunk: &mut Chunk, line: u32) {
    let active = chunk.add_import("web:html", "activeDocument");
    chunk.emit_call(active, 0, line);
}

fn child(chunk: &mut Chunk, parent: u16, first: bool, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, parent, line);
    let method = chunk.add_import("web:dom", if first { "firstChild" } else { "nextSibling" });
    chunk.emit_call(method, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn if_node(chunk: &mut Chunk, node: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    let truthy = chunk.add_import("ecma:boolean", "toBoolean");
    chunk.emit_call(truthy, 1, line);
    chunk.emit_if(line);
}

fn attribute_get(chunk: &mut Chunk, node: u16, name: &str, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const(name, line);
    let get = chunk.add_import("web:dom", "getAttribute");
    chunk.emit_call(get, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn attribute_set(chunk: &mut Chunk, node: u16, name: &str, value: u16, line: u32) {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const(name, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let set = chunk.add_import("web:dom", "setAttribute");
    chunk.emit_call(set, 4, line);
    chunk.emit_op(Op::DROP, line);
}

const TAB_STATES: &str = "__dotnet_tab_states";

fn tab_state(chunk: &mut Chunk, control: u16, line: u32) -> (u16, u16) {
    let registry = chunk.alloc_scratch(1);
    globals::emit_read(chunk, TAB_STATES, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let create_block = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);
    let new = chunk.add_import("ecma:object", "new");
    chunk.emit_call(new, 0, line);
    core_wasm::dup(chunk, line);
    globals::emit_write(chunk, TAB_STATES, line);
    chunk.emit_end(line);
    chunk.patch_block(create_block);
    chunk.emit_op_u16(Op::LOCAL_SET, registry, line);

    let id = attribute_get(chunk, control, "id", line);
    let first = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, registry, line);
    chunk.emit_op_u16(Op::LOCAL_GET, id, line);
    let has = chunk.add_import("ecma:object", "hasOwn");
    chunk.emit_call(has, 2, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op_u16(Op::LOCAL_SET, first, line);
    let state = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, first, line);
    chunk.emit_if(line);
    let new = chunk.add_import("ecma:object", "new");
    chunk.emit_call(new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, state, line);
    chunk.emit_op_u16(Op::LOCAL_GET, registry, line);
    chunk.emit_op_u16(Op::LOCAL_GET, id, line);
    chunk.emit_op_u16(Op::LOCAL_GET, state, line);
    let set = chunk.add_import("ecma:object", "set");
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, registry, line);
    chunk.emit_op_u16(Op::LOCAL_GET, id, line);
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_call(get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, state, line);
    chunk.emit_end(line);
    (state, first)
}
fn style(chunk: &mut Chunk, node: u16, property: &str, value: &str, line: u32) {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const(property, line);
    chunk.emit_string_const(value, line);
    let set = chunk.add_import("web:cssom", "setStyleProperty");
    chunk.emit_call(set, 4, line);
    chunk.emit_op(Op::DROP, line);
}

fn field_get(chunk: &mut Chunk, control: u16, name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const(name, line);
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_call(get, 2, line);
}

fn field_set(chunk: &mut Chunk, control: u16, name: &str, value: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const(name, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let set = chunk.add_import("ecma:object", "set");
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
}

fn select(chunk: &mut Chunk, page: u16, button: u16, line: u32) {
    if_node(chunk, page, line);
    if_node(chunk, button, line);
    globals::emit_read(chunk, TAB_STATES, line);
    let page_id = attribute_get(chunk, page, "id", line);
    chunk.emit_op_u16(Op::LOCAL_GET, page_id, line);
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_call(get, 2, line);
    let state = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, state, line);
    if_node(chunk, state, line);
    field_get(chunk, state, "page", line);
    let old_page = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, old_page, line);
    if_node(chunk, old_page, line);
    style(chunk, old_page, "display", "none", line);
    chunk.emit_end(line);
    field_get(chunk, state, "button", line);
    let old_button = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, old_button, line);
    if_node(chunk, old_button, line);
    style(chunk, old_button, "background-color", "#e5e5e5", line);
    style(chunk, old_button, "color", "#5b6470", line);
    style(chunk, old_button, "font-weight", "400", line);
    let false_value = chunk.alloc_scratch(1);
    chunk.emit_string_const("false", line);
    chunk.emit_op_u16(Op::LOCAL_SET, false_value, line);
    attribute_set(chunk, old_button, "aria-selected", false_value, line);
    chunk.emit_end(line);
    style(chunk, page, "display", "block", line);
    style(chunk, button, "background-color", "#ffffff", line);
    style(chunk, button, "color", "#111827", line);
    style(chunk, button, "font-weight", "600", line);
    let true_value = chunk.alloc_scratch(1);
    chunk.emit_string_const("true", line);
    chunk.emit_op_u16(Op::LOCAL_SET, true_value, line);
    attribute_set(chunk, button, "aria-selected", true_value, line);
    field_set(chunk, state, "page", page, line);
    field_set(chunk, state, "button", button, line);
    for _ in 0..3 {
        chunk.emit_end(line);
    }
}

fn click_callback(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut body =
        vybe_compiler::primitives::functions::create_function_chunk("__dotnet_tab_select", 1);
    body.local_count = 1;
    field_get(&mut body, 0, "currentTarget", line);
    let target = body.alloc_scratch(1);
    body.emit_op_u16(Op::LOCAL_SET, target, line);
    let new = body.add_import("ecma:object", "new");
    body.emit_call(new, 0, line);
    let button = body.alloc_scratch(1);
    body.emit_op_u16(Op::LOCAL_SET, button, line);
    field_set(&mut body, button, "__node", target, line);
    let page_id = attribute_get(&mut body, button, "data-tab-page-id", line);
    document(&mut body, line);
    body.emit_op_u16(Op::LOCAL_GET, page_id, line);
    let get_page = body.add_import("web:dom", "getElementById");
    body.emit_call(get_page, 2, line);
    let page = body.alloc_scratch(1);
    body.emit_op_u16(Op::LOCAL_SET, page, line);
    select(&mut body, page, button, line);
    body.emit_ref_null(HT_EXTERN, line);
    body.emit_op(Op::RETURN, line);
    chunks.push(body);
    chunks.len() - 1
}

/// Stack: [TabControl, TabPage] -> [null].
pub fn emit_add_page(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    let page = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, page, line);
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    let strip = child(c, control, true, line);
    let pages = child(c, strip, false, line);

    let (state, first) = tab_state(c, control, line);

    document(c, line);
    c.emit_string_const("button", line);
    c.emit_string_const("", line);
    let create = c.add_import("web:dom", "createElement");
    c.emit_call(create, 3, line);
    let button = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, button, line);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, button, line);
    c.emit_string_const("type", line);
    c.emit_string_const("button", line);
    let attribute = c.add_import("web:dom", "setAttribute");
    c.emit_call(attribute, 4, line);
    c.emit_op(Op::DROP, line);

    let page_id = attribute_get(c, page, "id", line);
    attribute_set(c, button, "data-tab-page-id", page_id, line);
    globals::emit_read(c, TAB_STATES, line);
    c.emit_op_u16(Op::LOCAL_GET, page_id, line);
    c.emit_op_u16(Op::LOCAL_GET, state, line);
    let set = c.add_import("ecma:object", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
    c.emit_op_u16(Op::LOCAL_GET, page_id, line);
    c.emit_string_const("-tab", line);
    ops::emit_dyn_add(c, line);
    let button_id = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, button_id, line);
    attribute_set(c, button, "id", button_id, line);
    let selected = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, first, line);
    c.emit_if_value(line);
    c.emit_string_const("true", line);
    c.emit_else(line);
    c.emit_string_const("false", line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_SET, selected, line);
    attribute_set(c, button, "aria-selected", selected, line);

    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, button, line);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, page, line);
    c.emit_string_const("data-text", line);
    let text = c.add_import("web:dom", "getAttribute");
    c.emit_call(text, 3, line);
    let set_text = c.add_import("web:dom", "setTextContent");
    c.emit_call(set_text, 3, line);
    c.emit_op(Op::DROP, line);
    field_set(c, page, "__dotnet_tab_button", button, line);
    for (property, value) in [
        ("border", "1px solid #b5b5b5"),
        ("border-bottom", "none"),
        ("border-radius", "0"),
        ("padding", "3px 9px"),
        ("background-color", "#e5e5e5"),
        ("color", "#5b6470"),
        ("font-weight", "400"),
        ("font", "inherit"),
        ("cursor", "default"),
    ] {
        style(c, button, property, value, line);
    }
    let append = c.add_import("web:dom", "appendChild");
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, strip, line);
    c.emit_op_u16(Op::LOCAL_GET, button, line);
    c.emit_call(append, 3, line);
    c.emit_op(Op::DROP, line);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, pages, line);
    c.emit_op_u16(Op::LOCAL_GET, page, line);
    c.emit_call(append, 3, line);
    c.emit_op(Op::DROP, line);

    c.emit_op_u16(Op::LOCAL_GET, first, line);
    c.emit_if(line);
    style(c, page, "display", "block", line);
    style(c, button, "background-color", "#ffffff", line);
    style(c, button, "color", "#111827", line);
    style(c, button, "font-weight", "600", line);
    c.emit_else(line);
    style(c, page, "display", "none", line);
    c.emit_end(line);

    c.emit_op_u16(Op::LOCAL_GET, first, line);
    c.emit_if(line);
    field_set(c, state, "page", page, line);
    field_set(c, state, "button", button, line);
    c.emit_end(line);

    let callback = click_callback(chunks, line);
    let c = &mut chunks[current];
    vybe_compiler::primitives::gui::emit_add_event_listener(c, button, "click", callback, line);
    c.emit_ref_null(HT_EXTERN, line);
}
