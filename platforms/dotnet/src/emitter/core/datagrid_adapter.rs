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
//! ⚠ A `<th>` is appended to the TABLE, not to a header row this builds. CSS
//! 2.1 §17.2.1 generates an ANONYMOUS row box around cells that sit directly
//! in a table, so the header row exists without anyone declaring it — which is
//! the whole reason these adapters need no find-or-create and therefore no
//! branching. `widgets` implements that rule
//! (`flow_layout::ANONYMOUS_ROW`); without it every `Columns.Add` would render
//! nothing.
//!
//! Every host call pushes exactly one value, so each is followed by a `DROP`
//! except the one whose result the member call itself yields.

use vybe_compiler::primitives::gui::{CSSOM_MODULE, DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT};
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
fn create_into(chunks: &mut [Chunk], current: usize, slot: u16, tag: &str, line: u32) {
    document(chunks, current, line);
    chunks[current].emit_string_const(tag, line);
    // `createElement`'s second argument is the input TYPE, which only an
    // `<input>` has. A cell has none, and passing empty is what every other
    // caller of this signature does.
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "createElement", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
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
fn style_cell(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    for (property, value) in [("border", "1px solid #c8c8c8"), ("padding", "3px 6px")] {
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
fn set_text(chunks: &mut [Chunk], current: usize, node_slot: u16, text_slot: u16, line: u32) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    call_import(chunks, current, "setTextContent", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `parent.appendChild(child)`, result dropped.
fn append(chunks: &mut [Chunk], current: usize, parent_slot: u16, child_slot: u16, line: u32) {
    document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_slot, line);
    call_import(chunks, current, "appendChild", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `grid.Columns.Add(headerText)` / `.Add(name, headerText)`.
///
/// WinForms' two-argument overload is `(name, headerText)`: the name is the
/// programmatic key and the header is what the user reads. Only the header is
/// rendered, so the name is dropped rather than written to an attribute that
/// nothing would read — an unmapped value stored somewhere plausible is the
/// silent-wrong the attribute fallback already teaches against.
///
/// Stack in `[grid, …]`, out `[column]` — the `<th>` itself, so a caller that
/// keeps the result has the element the column IS.
pub fn emit_add_column(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let (grid, header, cell) = (base, base + 1, base + 2);

    // The header text is the LAST argument in both overloads, so it is on top
    // whichever one was written.
    chunks[current].emit_op_u16(Op::LOCAL_SET, header, line);
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, grid, line);

    create_into(chunks, current, cell, "th", line);
    set_text(chunks, current, cell, header, line);
    style_cell(chunks, current, cell, line);
    append(chunks, current, grid, cell, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cell, line);
}

/// `grid.Rows.Add(v1, v2, …)` — one `<tr>`, one `<td>` per value.
///
/// Registered per ARITY, because a variadic call has no single shape in
/// bytecode: `overloads` keys on argument count and each count reaches this
/// with `argc` telling it how many cells to build.
///
/// Stack in `[grid, v1, …, vn]`, out `[row]`.
pub fn emit_add_row(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let values = argc.saturating_sub(1) as usize;
    let base = chunks[current].alloc_scratch(3 + values as u16);
    let (grid, row, cell) = (base, base + 1, base + 2);
    let value_base = base + 3;

    // Popped LAST-first, so the slots end up in written order.
    for index in (0..values).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_base + index as u16, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, grid, line);

    create_into(chunks, current, row, "tr", line);
    for index in 0..values {
        create_into(chunks, current, cell, "td", line);
        set_text(chunks, current, cell, value_base + index as u16, line);
        style_cell(chunks, current, cell, line);
        append(chunks, current, row, cell, line);
    }
    // The row joins the table LAST. Appending it first would make the table
    // re-measure its columns once per cell, and every one of those passes would
    // be against a row that was still being filled.
    append(chunks, current, grid, row, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row, line);
}
