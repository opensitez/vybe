//! DataGridView selection on its DOM rows and cells.

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::{loops, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

use super::datagrid_binding_adapter;
use super::object_fields::field_slot;

fn document(c: &mut Chunk, line: u32) {
    let f = c.add_import("web:html", "activeDocument");
    c.emit_call(f, 0, line);
}

pub fn emit_collection_item(c: &mut Chunk, line: u32) {
    c.emit_op(Op::ARRAY_GET, line);
}

fn dom(c: &mut Chunk, name: &str, argc: u8, line: u32) {
    let f = c.add_import("web:dom", name);
    c.emit_call(f, argc, line);
}

fn event_part(c: &mut Chunk, name: &str, line: u32) -> u16 {
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_string_const(name, line);
    let get = c.add_import("ecma:object", "get");
    c.emit_call(get, 2, line);
    let slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn attribute(c: &mut Chunk, node: u16, name: &str, line: u32) -> u16 {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(name, line);
    dom(c, "getAttribute", 3, line);
    let slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn set_attribute(c: &mut Chunk, node: u16, name: &str, value: u16, line: u32) {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(name, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    dom(c, "setAttribute", 4, line);
    c.emit_op(Op::DROP, line);
}

fn remove_attribute(c: &mut Chunk, node: u16, name: &str, line: u32) {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(name, line);
    dom(c, "removeAttribute", 3, line);
    c.emit_op(Op::DROP, line);
}

fn query(c: &mut Chunk, root: u16, selector: &str, all: bool, line: u32) -> u16 {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, root, line);
    c.emit_string_const(selector, line);
    dom(
        c,
        if all {
            "querySelectorAllWithin"
        } else {
            "querySelectorWithin"
        },
        3,
        line,
    );
    let slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn parent(c: &mut Chunk, node: u16, line: u32) -> u16 {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    dom(c, "parentNode", 2, line);
    let slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn style_row(c: &mut Chunk, row: u16, selected: bool, line: u32) {
    for (property, value) in [("background-color", "#0078d7"), ("color", "#fff")] {
        document(c, line);
        c.emit_op_u16(Op::LOCAL_GET, row, line);
        c.emit_string_const(property, line);
        if selected {
            c.emit_string_const(value, line);
            let set = c.add_import("web:cssom", "setStyleProperty");
            c.emit_call(set, 4, line);
        } else {
            let remove = c.add_import("web:cssom", "removeStyleProperty");
            c.emit_call(remove, 3, line);
        }
        c.emit_op(Op::DROP, line);
    }
}

fn selected_cell(c: &mut Chunk, row: u16, cell: u16, line: u32) -> u16 {
    let result = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, cell, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    let first = query(c, row, "td", false, line);
    c.emit_op_u16(Op::LOCAL_GET, first, line);
    c.emit_op_u16(Op::LOCAL_SET, result, line);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, cell, line);
    c.emit_op_u16(Op::LOCAL_SET, result, line);
    c.emit_end(line);
    result
}

pub(super) fn select(c: &mut Chunk, grid: u16, row: u16, cell: u16, line: u32) {
    for (selector, is_row) in [
        ("tbody tr[aria-selected=true]", true),
        ("tbody td[aria-selected=true]", false),
    ] {
        let selected = query(c, grid, selector, true, line);
        let index = c.alloc_scratch(1);
        c.emit_i32_const(0, line);
        c.emit_op_u16(Op::LOCAL_SET, index, line);
        let state = loops::emit_loop_start(std::slice::from_mut(c), 0, line);
        c.emit_op_u16(Op::LOCAL_GET, index, line);
        c.emit_op_u16(Op::LOCAL_GET, selected, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        c.emit_op(Op::I32_LT_S, line);
        loops::emit_loop_cond_from_i32(std::slice::from_mut(c), 0, line);
        c.emit_op_u16(Op::LOCAL_GET, selected, line);
        c.emit_op_u16(Op::LOCAL_GET, index, line);
        c.emit_op(Op::ARRAY_GET, line);
        let node = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, node, line);
        remove_attribute(c, node, "aria-selected", line);
        if is_row {
            style_row(c, node, false, line);
        }
        c.emit_op_u16(Op::LOCAL_GET, index, line);
        c.emit_i32_const(1, line);
        c.emit_op(Op::I32_ADD, line);
        c.emit_op_u16(Op::LOCAL_SET, index, line);
        loops::emit_loop_end(std::slice::from_mut(c), 0, state, line);
    }
    add_selection(c, row, cell, line);
}

fn add_selection(c: &mut Chunk, row: u16, cell: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, row, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    let true_value = c.alloc_scratch(1);
    c.emit_string_const("true", line);
    c.emit_op_u16(Op::LOCAL_SET, true_value, line);
    set_attribute(c, row, "aria-selected", true_value, line);
    style_row(c, row, true, line);
    let cell = selected_cell(c, row, cell, line);
    c.emit_op_u16(Op::LOCAL_GET, cell, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    set_attribute(c, cell, "aria-selected", true_value, line);
    c.emit_end(line);
    c.emit_end(line);
}

fn toggle_selection(c: &mut Chunk, row: u16, cell: u16, line: u32) {
    let selected = attribute(c, row, "aria-selected", line);
    c.emit_op_u16(Op::LOCAL_GET, selected, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    add_selection(c, row, cell, line);
    c.emit_else(line);
    remove_attribute(c, row, "aria-selected", line);
    style_row(c, row, false, line);
    let selected_cell = query(c, row, "td[aria-selected=true]", false, line);
    c.emit_op_u16(Op::LOCAL_GET, selected_cell, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    remove_attribute(c, selected_cell, "aria-selected", line);
    c.emit_end(line);
    c.emit_end(line);
}

pub(super) fn set_row_index(c: &mut Chunk, row: u16, index: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    strings::emit_to_string(c, line);
    let text = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, text, line);
    set_attribute(c, row, "data-row-index", text, line);
}

pub(super) fn set_column_index(c: &mut Chunk, cell: u16, index: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    strings::emit_to_string(c, line);
    let text = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, text, line);
    set_attribute(c, cell, "data-column-index", text, line);
}

pub(super) fn set_column_key(c: &mut Chunk, cell: u16, key: u16, line: u32) {
    set_attribute(c, cell, "data-column-key", key, line);
}

pub(super) fn select_at_position(c: &mut Chunk, grid: u16, source: u16, line: u32) {
    let position = c.alloc_scratch(1);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, position, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    datagrid_binding_adapter::is_binding_source(c, source, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    class_slots::emit_class_get(
        c,
        ObjSource::Stack,
        &field_slot("position"),
        Dest::Stack,
        line,
    );
    c.emit_op_u16(Op::LOCAL_SET, position, line);
    c.emit_end(line);
    c.emit_end(line);
    let rows = query(c, grid, "tbody tr", true, line);
    c.emit_op_u16(Op::LOCAL_GET, position, line);
    c.emit_op_u16(Op::LOCAL_GET, rows, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    ops::emit_dyn_lt(c, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, rows, line);
    c.emit_op_u16(Op::LOCAL_GET, position, line);
    c.emit_op(Op::ARRAY_GET, line);
    let row = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, row, line);
    c.emit_ref_null(HT_EXTERN, line);
    let no_cell = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, no_cell, line);
    select(c, grid, row, no_cell, line);
    c.emit_end(line);
}

fn callback(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut body = vec![vybe_compiler::primitives::functions::create_function_chunk(
        "__dotnet_grid_select",
        1,
    )];
    body[0].local_count = 1;
    let target = event_part(&mut body[0], "target", line);
    let grid = event_part(&mut body[0], "currentTarget", line);
    document(&mut body[0], line);
    body[0].emit_op_u16(Op::LOCAL_GET, target, line);
    body[0].emit_string_const("tr", line);
    dom(&mut body[0], "closest", 3, line);
    let row = body[0].alloc_scratch(1);
    body[0].emit_op_u16(Op::LOCAL_SET, row, line);
    body[0].emit_op_u16(Op::LOCAL_GET, row, line);
    body[0].emit_op(Op::REF_IS_NULL, line);
    body[0].emit_op(Op::I32_EQZ, line);
    body[0].emit_if(line);
    let row_index = attribute(&mut body[0], row, "data-row-index", line);
    body[0].emit_op_u16(Op::LOCAL_GET, row_index, line);
    body[0].emit_op(Op::REF_IS_NULL, line);
    body[0].emit_op(Op::I32_EQZ, line);
    body[0].emit_if(line);
    document(&mut body[0], line);
    body[0].emit_op_u16(Op::LOCAL_GET, target, line);
    body[0].emit_string_const("td", line);
    dom(&mut body[0], "closest", 3, line);
    let cell = body[0].alloc_scratch(1);
    body[0].emit_op_u16(Op::LOCAL_SET, cell, line);
    let ctrl = event_part(&mut body[0], "ctrlKey", line);
    let meta = event_part(&mut body[0], "metaKey", line);
    body[0].emit_op_u16(Op::LOCAL_GET, grid, line);
    body[0].emit_string_const("multiselect", line);
    let get = body[0].add_import("ecma:object", "get");
    body[0].emit_call(get, 2, line);
    ops::emit_dyn_to_bool(&mut body[0], line);
    body[0].emit_op_u16(Op::LOCAL_GET, ctrl, line);
    ops::emit_dyn_to_bool(&mut body[0], line);
    body[0].emit_op_u16(Op::LOCAL_GET, meta, line);
    ops::emit_dyn_to_bool(&mut body[0], line);
    body[0].emit_op(Op::I32_OR, line);
    body[0].emit_op(Op::I32_AND, line);
    body[0].emit_if(line);
    toggle_selection(&mut body[0], row, cell, line);
    body[0].emit_else(line);
    select(&mut body[0], grid, row, cell, line);
    body[0].emit_end(line);
    body[0].emit_op_u16(Op::LOCAL_GET, grid, line);
    datagrid_binding_adapter::get_source(&mut body[0], line);
    let source = body[0].alloc_scratch(1);
    body[0].emit_op_u16(Op::LOCAL_SET, source, line);
    body[0].emit_op_u16(Op::LOCAL_GET, source, line);
    ops::emit_dyn_to_bool(&mut body[0], line);
    body[0].emit_if(line);
    datagrid_binding_adapter::is_binding_source(&mut body[0], source, line);
    body[0].emit_if(line);
    body[0].emit_op_u16(Op::LOCAL_GET, source, line);
    body[0].emit_op_u16(Op::LOCAL_GET, row_index, line);
    let number = body[0].add_import("ecma:number", "Number");
    body[0].emit_call(number, 1, line);
    class_slots::emit_class_set(
        &mut body[0],
        ObjSource::Stack,
        &field_slot("position"),
        ValueSource::Stack,
        line,
    );
    super::control_binding_adapter::refresh_controls(&mut body, 0, source, line);
    super::bindingnavigator_adapter::refresh_source(&mut body, 0, source, line);
    body[0].emit_end(line);
    body[0].emit_end(line);
    body[0].emit_end(line);
    body[0].emit_end(line);
    body[0].emit_ref_null(HT_EXTERN, line);
    body[0].emit_op(Op::RETURN, line);
    chunks.push(body.remove(0));
    chunks.len() - 1
}

pub(super) fn register(chunks: &mut Vec<Chunk>, current: usize, grid: u16, line: u32) {
    let marker = attribute(
        &mut chunks[current],
        grid,
        "data-vybe-grid-select-listener",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, marker, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    let handler = callback(chunks, line);
    vybe_compiler::primitives::gui::emit_add_event_listener(
        &mut chunks[current],
        grid,
        "click",
        handler,
        line,
    );
    let true_value = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("true", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, true_value, line);
    set_attribute(
        &mut chunks[current],
        grid,
        "data-vybe-grid-select-listener",
        true_value,
        line,
    );
    chunks[current].emit_end(line);
}

pub fn emit_current_row(c: &mut Chunk, line: u32) {
    let grid = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, grid, line);
    let row = query(c, grid, "tbody tr[aria-selected=true]", false, line);
    c.emit_op_u16(Op::LOCAL_GET, row, line);
}

pub fn emit_selected_rows(c: &mut Chunk, line: u32) {
    let grid = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, grid, line);
    let rows = query(c, grid, "tbody tr[aria-selected=true]", true, line);
    c.emit_op_u16(Op::LOCAL_GET, rows, line);
}

pub fn emit_current_cell(c: &mut Chunk, line: u32) {
    let grid = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, grid, line);
    let cell = query(c, grid, "tbody td[aria-selected=true]", false, line);
    c.emit_op_u16(Op::LOCAL_GET, cell, line);
}

pub fn emit_selected_cells(c: &mut Chunk, line: u32) {
    let grid = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, grid, line);
    let cells = query(c, grid, "tbody td[aria-selected=true]", true, line);
    c.emit_op_u16(Op::LOCAL_GET, cells, line);
}

pub fn emit_selection_count(c: &mut Chunk, line: u32) {
    c.emit_op(Op::ARRAY_LENGTH, line);
}

pub fn emit_row_cells(c: &mut Chunk, line: u32) {
    let row = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, row, line);
    let cells = query(c, row, "td", true, line);
    c.emit_op_u16(Op::LOCAL_GET, cells, line);
}

pub fn emit_cell_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let cell = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, cell, line);
    let key = attribute(c, cell, "data-column-key", line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if_value(line);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, cell, line);
    dom(c, "textContent", 2, line);
    c.emit_else(line);
    let row = parent(c, cell, line);
    c.emit_op_u16(Op::LOCAL_GET, row, line);
    emit_bound_item(chunks, current, line);
    let c = &mut chunks[current];
    let item = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, item, line);
    c.emit_op_u16(Op::LOCAL_GET, item, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, item, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    let get = c.add_import("ecma:object", "get");
    c.emit_call(get, 2, line);
    c.emit_else(line);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, cell, line);
    dom(c, "textContent", 2, line);
    c.emit_end(line);
    c.emit_end(line);
}

pub fn emit_row_index(c: &mut Chunk, line: u32) {
    let row = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, row, line);
    let index = attribute(c, row, "data-row-index", line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    let number = c.add_import("ecma:number", "Number");
    c.emit_call(number, 1, line);
}

pub fn emit_column_index(c: &mut Chunk, line: u32) {
    let cell = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, cell, line);
    let index = attribute(c, cell, "data-column-index", line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    let number = c.add_import("ecma:number", "Number");
    c.emit_call(number, 1, line);
}

pub fn emit_bound_item(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let row = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, row, line);
    let index_text = attribute(c, row, "data-row-index", line);
    let body = parent(c, row, line);
    let table = parent(c, body, line);
    let grid = parent(c, table, line);
    c.emit_op_u16(Op::LOCAL_GET, grid, line);
    datagrid_binding_adapter::get_source(c, line);
    let source = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, source, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if_value(line);
    c.emit_ref_null(HT_EXTERN, line);
    c.emit_else(line);
    let data = datagrid_binding_adapter::source_data(c, source, line);
    let (rows, _) = datagrid_binding_adapter::data_arrays(chunks, current, data, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, index_text, line);
    let number = c.add_import("ecma:number", "Number");
    c.emit_call(number, 1, line);
    let index = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op_u16(Op::LOCAL_GET, rows, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    ops::emit_dyn_lt(c, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, rows, line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_else(line);
    c.emit_ref_null(HT_EXTERN, line);
    c.emit_end(line);
    c.emit_end(line);
}
