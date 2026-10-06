//! `DataGridView.Columns` and `.Rows` — the grid's data surface, as DOM.
//!
//! **A grid's columns and rows are ELEMENTS, not a control's internal state.**
//! `Columns.Add("Name", "Name")` is a `<th>` appended to the table and
//! `Rows.Add(a, b)` is a `<tr>` of `<td>`s — which is what a real engine would
//! render and what `widgets`' CSS tables already lay out. Nothing here
//! adds a host function: the
//! adapters compose `web:html`'s `createElement` / `setTextContent` /
//! `appendChild`, the same three calls a script would make.
//!
//! The constructor supplies a real thead/tbody, so columns and rows stay in
//! their semantic table sections in both browser engines.
//!
//! Every host call pushes exactly one value, so each is followed by a `DROP`
//! except the one whose result the member call itself yields.

use vybe_compiler::primitives::gui::{CSSOM_MODULE, DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT};
use vybe_compiler::primitives::ops;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// The node operations live in `web:dom`; `activeDocument` is `web:html`'s.
/// Two modules, because the DOCUMENT is what the user agent hands you and the
/// nodes are what the DOM does with it.
const DOM_MODULE: &str = "web:dom";

fn call_import(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(DOM_MODULE, name);
    chunks[current].emit_call(idx, argc, line);
}

/// Push the active document — the first argument of every `web:html` call.
fn document(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
    chunks[current].emit_call(idx, 0, line);
}

/// `createElement(tag)` into `slot`. Leaves the stack as it found it.
pub(super) fn create_into(chunks: &mut [Chunk], current: usize, slot: u16, tag: &str, line: u32) {
    document(chunks, current, line);
    chunks[current].emit_string_const(tag, line);
    // `createElement`'s second argument is the input TYPE, which only an
    // `<input>` has. A cell has none, and passing empty is what every other
    // caller of this signature does.
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "createElement", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

pub(super) fn child_into(
    chunks: &mut [Chunk],
    current: usize,
    parent: u16,
    child: u16,
    first: bool,
    line: u32,
) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent, line);
    call_import(
        chunks,
        current,
        if first { "firstChild" } else { "lastChild" },
        2,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, child, line);
}

/// The control is the scroll viewport; its first child is the table.
pub(super) fn table_into(chunks: &mut [Chunk], current: usize, grid: u16, table: u16, line: u32) {
    child_into(chunks, current, grid, table, true, line);
}

/// The cell styling a `DataGridView` draws and a `<table>` does not.
///
/// **This is the CONTROL's default appearance, not HTML's.** A browser draws no
/// rules between table cells — correctly, because a table is a layout — while a
/// WinForms grid draws them and pads its cells. That difference belongs to the
/// control, so it is written here, on the cells the control builds, rather than
/// pushed into the UA sheet where it would restyle every `<table>` on the page.
///
/// It goes through `setStyleProperty`, so it cascades and serialises into the
/// `style` attribute like any author declaration — an author or a later
/// `GridColor` write overrides it exactly as a browser would let them.
pub(super) fn style_cell(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    for (property, value) in [
        ("border", "1px solid #c8c8c8"),
        ("padding", "3px 6px"),
        ("overflow", "hidden"),
        ("white-space", "nowrap"),
    ] {
        document(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_string_const(property, line);
        chunks[current].emit_string_const(value, line);
        let idx = chunks[current].add_import(CSSOM_MODULE, "setStyleProperty");
        chunks[current].emit_call(idx, 4, line);
        chunks[current].emit_op(Op::DROP, line);
    }
}

/// `node.textContent = <whatever is in `text_slot`>`.
pub(super) fn set_text(
    chunks: &mut [Chunk],
    current: usize,
    node_slot: u16,
    text_slot: u16,
    line: u32,
) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    call_import(chunks, current, "setTextContent", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `parent.appendChild(child)`, result dropped.
pub(super) fn append(
    chunks: &mut [Chunk],
    current: usize,
    parent_slot: u16,
    child_slot: u16,
    line: u32,
) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_slot, line);
    call_import(chunks, current, "appendChild", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `grid.Columns.Add(headerText)` / `.Add(name, headerText)`.
///
/// WinForms' two-argument overload is `(name, headerText)`: the name is the
/// programmatic key and the header is what the user reads. The binding
/// renderer uses the key to populate manually declared columns.
///
/// Stack in `[grid, …]`, out `[column]` — the `<th>` itself, so a caller that
/// keeps the result has the element the column IS.
pub fn emit_add_column(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(8);
    let (grid, header, cell, head, row, table, key, source) =
        (base, base + 1, base + 2, base + 3, base + 4, base + 5, base + 6, base + 7);

    // The header text is the LAST argument in both overloads, so it is on top
    // whichever one was written.
    chunks[current].emit_op_u16(Op::LOCAL_SET, header, line);
    if argc > 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, header, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, grid, line);

    table_into(chunks, current, grid, table, line);
    child_into(chunks, current, table, head, true, line);
    child_into(chunks, current, head, row, true, line);
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row, line);
    call_import(chunks, current, "children", 2, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    create_into(chunks, current, cell, "th", line);
    style_cell(chunks, current, cell, line);
    set_cell_width(chunks, current, cell, "24px", line);
    super::datagrid_properties_adapter::style_new_row_header(
        &mut chunks[current], grid, cell, line,
    );
    append(chunks, current, row, cell, line);
    chunks[current].emit_end(line);
    create_into(chunks, current, cell, "th", line);
    set_text(chunks, current, cell, header, line);
    style_cell(chunks, current, cell, line);
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cell, line);
    chunks[current].emit_string_const("data-grid-manual-key", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_import(chunks, current, "setAttribute", 4, line);
    chunks[current].emit_op(Op::DROP, line);
    append(chunks, current, row, cell, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, grid, line);
    super::datagrid_binding_adapter::get_source(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    super::datagrid_binding_adapter::render(chunks, current, grid, source, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cell, line);
}

/// `grid.Rows.Add(v1, v2, …)` — one `<tr>`, one `<td>` per value.
///
/// Registered per ARITY, because a variadic call has no single shape in
/// bytecode: `overloads` keys on argument count and each count reaches this
/// with `argc` telling it how many cells to build.
///
/// Stack in `[grid, v1, …, vn]`, out `[row]`.
pub fn emit_add_row(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let values = argc.saturating_sub(1) as usize;
    let base = chunks[current].alloc_scratch(5 + values as u16);
    let (grid, row, cell, body, table) = (base, base + 1, base + 2, base + 3, base + 4);
    let value_base = base + 5;

    // Popped LAST-first, so the slots end up in written order.
    for index in (0..values).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_base + index as u16, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, grid, line);

    table_into(chunks, current, grid, table, line);
    child_into(chunks, current, table, body, false, line);
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, body, line);
    call_import(chunks, current, "children", 2, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    let row_index = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, row_index, line);
    create_into(chunks, current, row, "tr", line);
    super::datagrid_selection_adapter::set_row_index(&mut chunks[current], row, row_index, line);
    create_into(chunks, current, cell, "th", line);
    style_cell(chunks, current, cell, line);
    set_cell_width(chunks, current, cell, "24px", line);
    super::datagrid_properties_adapter::style_new_row_header(
        &mut chunks[current], grid, cell, line,
    );
    append(chunks, current, row, cell, line);
    for index in 0..values {
        create_into(chunks, current, cell, "td", line);
        chunks[current].emit_i32_const(index as i32, line);
        let column_index = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, column_index, line);
        super::datagrid_selection_adapter::set_column_index(
            &mut chunks[current],
            cell,
            column_index,
            line,
        );
        set_text(chunks, current, cell, value_base + index as u16, line);
        style_cell(chunks, current, cell, line);
        append(chunks, current, row, cell, line);
    }
    // The row joins the table LAST. Appending it first would make the table
    // re-measure its columns once per cell, and every one of those passes would
    // be against a row that was still being filled.
    append(chunks, current, body, row, line);
    super::datagrid_selection_adapter::register(chunks, current, grid, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row_index, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    let no_cell = chunks[current].alloc_scratch(1);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, no_cell, line);
    super::datagrid_selection_adapter::select(&mut chunks[current], grid, row, no_cell, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row, line);
}

/// The grid header begins with a blank row-header corner; only later cells
/// represent DataGridView columns.
pub fn emit_column_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let (grid, head, row, count, table) = (base, base + 1, base + 2, base + 3, base + 4);
    chunks[current].emit_op_u16(Op::LOCAL_SET, grid, line);
    table_into(chunks, current, grid, table, line);
    child_into(chunks, current, table, head, true, line);
    child_into(chunks, current, head, row, true, line);
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row, line);
    call_import(chunks, current, "children", 2, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
}

pub(super) fn set_cell_width(
    chunks: &mut [Chunk],
    current: usize,
    cell: u16,
    width: &str,
    line: u32,
) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cell, line);
    chunks[current].emit_string_const("width", line);
    chunks[current].emit_string_const(width, line);
    let idx = chunks[current].add_import(CSSOM_MODULE, "setStyleProperty");
    chunks[current].emit_call(idx, 4, line);
    chunks[current].emit_op(Op::DROP, line);
}

pub fn emit_row_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let (grid, table, body) = (base, base + 1, base + 2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, grid, line);
    table_into(chunks, current, grid, table, line);
    child_into(chunks, current, table, body, false, line);
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, body, line);
    call_import(chunks, current, "children", 2, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
}
