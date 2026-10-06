//! DataGridView presentation properties on the grid's table elements.

use vybe_compiler::primitives::{loops, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

fn document(c: &mut Chunk, line: u32) {
    let f = c.add_import("web:html", "activeDocument");
    c.emit_call(f, 0, line);
}

fn property_get(c: &mut Chunk, grid: u16, name: &str, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, grid, line);
    c.emit_string_const(name, line);
    let f = c.add_import("ecma:object", "get");
    c.emit_call(f, 2, line);
}

fn property_set(c: &mut Chunk, grid: u16, name: &str, value: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, grid, line);
    c.emit_string_const(name, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let f = c.add_import("ecma:object", "set");
    c.emit_call(f, 3, line);
    c.emit_op(Op::DROP, line);
}

fn style(c: &mut Chunk, node: u16, property: &str, value: &str, line: u32) {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(property, line);
    c.emit_string_const(value, line);
    let f = c.add_import("web:cssom", "setStyleProperty");
    c.emit_call(f, 4, line);
    c.emit_op(Op::DROP, line);
}

fn row_header_display(c: &mut Chunk, visible: u16, header: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, visible, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    style(c, header, "display", "table-cell", line);
    c.emit_else(line);
    style(c, header, "display", "none", line);
    c.emit_end(line);
}

/// Called when either binding or Rows.Add creates a row-header cell.
pub(super) fn style_new_row_header(c: &mut Chunk, grid: u16, header: u16, line: u32) {
    property_get(c, grid, "rowheadersvisible", line);
    let visible = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, visible, line);
    row_header_display(c, visible, header, line);
}

/// Stack: [grid] -> [null].
pub fn emit_init(c: &mut Chunk, line: u32) {
    let grid = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, grid, line);
    for name in [
        "autogeneratecolumns",
        "columnheadersvisible",
        "multiselect",
        "rowheadersvisible",
    ] {
        c.emit_bool_const(true, line);
        let value = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, value, line);
        property_set(c, grid, name, value, line);
    }
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [grid] -> [Boolean].
pub fn emit_get(c: &mut Chunk, name: &str, line: u32) {
    let grid = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, grid, line);
    property_get(c, grid, name, line);
}

/// Stack: [grid, Boolean] -> [null].
pub fn emit_set(chunks: &mut Vec<Chunk>, current: usize, name: &str, line: u32) {
    let c = &mut chunks[current];
    let value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    let grid = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, grid, line);
    property_set(c, grid, name, value, line);

    if name == "multiselect" {
        c.emit_op_u16(Op::LOCAL_GET, value, line);
        ops::emit_dyn_to_bool(c, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_if(line);
        c.emit_op_u16(Op::LOCAL_GET, grid, line);
        super::datagrid_selection_adapter::emit_selected_rows(c, line);
        let rows = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, rows, line);
        c.emit_op_u16(Op::LOCAL_GET, rows, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        c.emit_i32_const(0, line);
        c.emit_op(Op::I32_GT_S, line);
        c.emit_if(line);
        c.emit_op_u16(Op::LOCAL_GET, rows, line);
        c.emit_i32_const(0, line);
        c.emit_op(Op::ARRAY_GET, line);
        let first = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, first, line);
        let no_cell = c.alloc_scratch(1);
        c.emit_ref_null(HT_EXTERN, line);
        c.emit_op_u16(Op::LOCAL_SET, no_cell, line);
        super::datagrid_selection_adapter::select(c, grid, first, no_cell, line);
        c.emit_end(line);
        c.emit_end(line);
        c.emit_ref_null(HT_EXTERN, line);
        return;
    }

    if name == "autogeneratecolumns" {
        c.emit_op_u16(Op::LOCAL_GET, grid, line);
        super::datagrid_binding_adapter::get_source(c, line);
        let source = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, source, line);
        c.emit_op_u16(Op::LOCAL_GET, source, line);
        ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        super::datagrid_binding_adapter::render(chunks, current, grid, source, line);
        chunks[current].emit_end(line);
        chunks[current].emit_ref_null(HT_EXTERN, line);
        return;
    }

    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, grid, line);
    c.emit_string_const(if name == "columnheadersvisible" {
        "thead"
    } else {
        "th:first-child"
    }, line);
    let query = c.add_import("web:dom", "querySelectorAllWithin");
    c.emit_call(query, 3, line);
    let nodes = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, nodes, line);
    let index = c.alloc_scratch(1);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op_u16(Op::LOCAL_GET, nodes, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, nodes, line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op(Op::ARRAY_GET, line);
    let node = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, node, line);
    if name == "columnheadersvisible" {
        c.emit_op_u16(Op::LOCAL_GET, value, line);
        ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        style(c, node, "display", "table-header-group", line);
        c.emit_else(line);
        style(c, node, "display", "none", line);
        c.emit_end(line);
    } else {
        row_header_display(c, value, node, line);
    }
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}
