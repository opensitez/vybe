//! DataGridView.DataSource rendered into the same table used by Columns/Rows.

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::{collections, globals, loops, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

use super::datagrid_adapter;
use super::object_fields::field_slot;

const GRID_SOURCES: &str = "__dotnet_grid_sources";

fn document(chunk: &mut Chunk, line: u32) {
    let active = chunk.add_import("web:html", "activeDocument");
    chunk.emit_call(active, 0, line);
}

fn get(chunk: &mut Chunk, object: u16, property: &str, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, object, line);
    chunk.emit_string_const(property, line);
    let method = chunk.add_import("ecma:object", "get");
    chunk.emit_call(method, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn is_binding_source(chunk: &mut Chunk, source: u16, line: u32) {
    let ty = get(chunk, source, "__type", line);
    chunk.emit_op_u16(Op::LOCAL_GET, ty, line);
    chunk.emit_string_const("BindingSource", line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ty, line);
    chunk.emit_string_const("System.Windows.Forms.BindingSource", line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn control_id(chunk: &mut Chunk, control: u16, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const("id", line);
    let method = chunk.add_import("web:dom", "getAttribute");
    chunk.emit_call(method, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn registry(chunk: &mut Chunk, line: u32) {
    globals::emit_read(chunk, GRID_SOURCES, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let block = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);
    let new = chunk.add_import("ecma:object", "new");
    chunk.emit_call(new, 0, line);
    core_wasm::dup(chunk, line);
    globals::emit_write(chunk, GRID_SOURCES, line);
    chunk.emit_end(line);
    chunk.patch_block(block);
}

fn empty_array(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    collections::emit_array_new(chunks, current, 0, line);
    let slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn source_data(chunk: &mut Chunk, source: u16, line: u32) -> u16 {
    is_binding_source(chunk, source, line);
    let data = chunk.alloc_scratch(1);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot("datasource"),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, data, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    chunk.emit_op_u16(Op::LOCAL_SET, data, line);
    chunk.emit_end(line);
    data
}

fn data_arrays(chunks: &mut [Chunk], current: usize, data: u16, line: u32) -> (u16, u16) {
    let rows = chunks[current].alloc_scratch(1);
    let columns = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    let empty_rows = empty_array(chunks, current, line);
    let empty_columns = empty_array(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, empty_rows, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rows, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, empty_columns, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, columns, line);
    chunks[current].emit_else(line);
    let table_rows = get(&mut chunks[current], data, "rows", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, table_rows, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rows, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rows, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rows, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    let keys = chunks[current].add_import("ecma:object", "keys");
    chunks[current].emit_call(keys, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, columns, line);
    chunks[current].emit_else(line);
    let empty = empty_array(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, empty, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, columns, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    let empty_rows = empty_array(chunks, current, line);
    let empty_columns = empty_array(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, empty_rows, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rows, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, empty_columns, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, columns, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, table_rows, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rows, line);
    let table_columns = get(&mut chunks[current], data, "columns", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, table_columns, line);
    chunks[current].emit_call(is_array, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    let empty = empty_array(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, empty, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, columns, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, table_columns, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, columns, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    (rows, columns)
}

fn clear_node(chunk: &mut Chunk, node: u16, line: u32) {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const("", line);
    let set = chunk.add_import("web:dom", "setTextContent");
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
}

fn array_len(chunk: &mut Chunk, array: u16, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, array, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn array_item(chunk: &mut Chunk, array: u16, index: u16, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, array, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn next(chunk: &mut Chunk, index: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
}

/// Render a DataTable or an array of objects into a DataGridView table.
pub fn render(chunks: &mut [Chunk], current: usize, grid: u16, source: u16, line: u32) {
    let head = chunks[current].alloc_scratch(1);
    let head_row = chunks[current].alloc_scratch(1);
    let body = chunks[current].alloc_scratch(1);
    datagrid_adapter::child_into(chunks, current, grid, head, true, line);
    datagrid_adapter::child_into(chunks, current, head, head_row, true, line);
    datagrid_adapter::child_into(chunks, current, grid, body, false, line);
    clear_node(&mut chunks[current], head_row, line);
    clear_node(&mut chunks[current], body, line);
    let data = source_data(&mut chunks[current], source, line);
    let (rows, columns) = data_arrays(chunks, current, data, line);
    let column_count = array_len(&mut chunks[current], columns, line);
    let row_count = array_len(&mut chunks[current], rows, line);

    let cell = chunks[current].alloc_scratch(1);
    datagrid_adapter::create_into(chunks, current, cell, "th", line);
    datagrid_adapter::style_cell(chunks, current, cell, line);
    datagrid_adapter::append(chunks, current, head_row, cell, line);

    let column_index = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, column_index, line);
    let column_loop = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, column_index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, column_count, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    let column = array_item(&mut chunks[current], columns, column_index, line);
    datagrid_adapter::create_into(chunks, current, cell, "th", line);
    datagrid_adapter::set_text(chunks, current, cell, column, line);
    datagrid_adapter::style_cell(chunks, current, cell, line);
    datagrid_adapter::append(chunks, current, head_row, cell, line);
    next(&mut chunks[current], column_index, line);
    loops::emit_loop_end(chunks, current, column_loop, line);

    let row_index = chunks[current].alloc_scratch(1);
    let tr = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, row_index, line);
    let row_loop = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row_index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row_count, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    let row = array_item(&mut chunks[current], rows, row_index, line);
    datagrid_adapter::create_into(chunks, current, tr, "tr", line);
    datagrid_adapter::create_into(chunks, current, cell, "th", line);
    datagrid_adapter::style_cell(chunks, current, cell, line);
    datagrid_adapter::append(chunks, current, tr, cell, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, column_index, line);
    let value_loop = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, column_index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, column_count, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    let column = array_item(&mut chunks[current], columns, column_index, line);
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, column, line);
    let get = chunks[current].add_import("ecma:object", "get");
    chunks[current].emit_call(get, 2, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, column, line);
    chunks[current].emit_call(get, 2, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    datagrid_adapter::create_into(chunks, current, cell, "td", line);
    datagrid_adapter::set_text(chunks, current, cell, value, line);
    datagrid_adapter::style_cell(chunks, current, cell, line);
    datagrid_adapter::append(chunks, current, tr, cell, line);
    next(&mut chunks[current], column_index, line);
    loops::emit_loop_end(chunks, current, value_loop, line);
    datagrid_adapter::append(chunks, current, body, tr, line);
    next(&mut chunks[current], row_index, line);
    loops::emit_loop_end(chunks, current, row_loop, line);
}

/// Stack: [grid, source] -> [null].
pub fn set_source(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let source = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, source, line);
    let grid = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, grid, line);
    registry(c, line);
    let sources = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, sources, line);
    let id = control_id(c, grid, line);
    c.emit_op_u16(Op::LOCAL_GET, sources, line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    let set = c.add_import("ecma:object", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
    is_binding_source(c, source, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    class_slots::emit_class_get(
        c,
        ObjSource::Stack,
        &field_slot("__grids"),
        Dest::Stack,
        line,
    );
    c.emit_op_u16(Op::LOCAL_GET, grid, line);
    let push = c.add_import("ecma:array", "push");
    c.emit_call(push, 2, line);
    c.emit_op(Op::DROP, line);
    c.emit_end(line);
    render(chunks, current, grid, source, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// Stack: [grid] -> [source].
pub fn get_source(chunk: &mut Chunk, line: u32) {
    let grid = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, grid, line);
    let id = control_id(chunk, grid, line);
    registry(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, id, line);
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_call(get, 2, line);
}

pub fn refresh_source_grids(chunks: &mut [Chunk], current: usize, source: u16, line: u32) {
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    class_slots::emit_class_get(
        c,
        ObjSource::Stack,
        &field_slot("__grids"),
        Dest::Stack,
        line,
    );
    let grids = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, grids, line);
    let count = array_len(c, grids, line);
    let index = c.alloc_scratch(1);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    let grid = array_item(&mut chunks[current], grids, index, line);
    let id = control_id(&mut chunks[current], grid, line);
    globals::emit_read(&mut chunks[current], GRID_SOURCES, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, id, line);
    let get = chunks[current].add_import("ecma:object", "get");
    chunks[current].emit_call(get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    render(chunks, current, grid, source, line);
    chunks[current].emit_end(line);
    next(&mut chunks[current], index, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
}
