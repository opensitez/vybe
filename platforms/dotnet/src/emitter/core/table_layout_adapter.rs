//! WinForms table-layout styles and cell coordinates backed by CSS Grid.

use vybe_compiler::primitives::class_slots::{self, ClassSlot, Dest, ObjSource, PlainNames};
use vybe_compiler::primitives::{gui, ops, strings};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

fn field(chunk: &mut Chunk, object: u16, name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, object, line);
    let slot = class_slots::resolve(&ClassSlot::internal(name), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &slot, Dest::Stack, line);
}

fn document(chunk: &mut Chunk, line: u32) {
    let active = chunk.add_import(gui::DOCUMENT_MODULE, gui::HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(active, 0, line);
}

fn axis_name(chunk: &mut Chunk, axis: u16, columns: &str, rows: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, axis, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const(columns, line);
    chunk.emit_else(line);
    chunk.emit_string_const(rows, line);
    chunk.emit_end(line);
}

/// Stack: `[panel] -> [style collection]`. The panel owns the state, so
/// repeated `ColumnStyles` reads see the same entries.
pub fn emit_styles_get(chunk: &mut Chunk, row: bool, line: u32) {
    chunk.emit_i32_const(i32::from(row), line);
    crate::emitter::dispatch::emit_value_type_new(
        chunk,
        "TableLayoutStyleCollection",
        &["owner", "axis"],
        line,
    );
}

/// Stack: `[collection] -> [count]`.
pub fn emit_styles_count(chunk: &mut Chunk, line: u32) {
    let collection = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, collection, line);
    let owner = chunk.alloc_scratch(1);
    field(chunk, collection, "owner", line);
    chunk.emit_op_u16(Op::LOCAL_SET, owner, line);
    let axis = chunk.alloc_scratch(1);
    field(chunk, collection, "axis", line);
    chunk.emit_op_u16(Op::LOCAL_SET, axis, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, owner, line);
    axis_name(chunk, axis, "data-column-style-count", "data-row-style-count", line);
    let getter = chunk.add_import("web:dom", "getAttribute");
    chunk.emit_call(getter, 3, line);
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
}

/// Stack: `[collection, style] -> [index]`.
pub fn emit_style_add(chunk: &mut Chunk, line: u32) {
    let style = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, style, line);
    let collection = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, collection, line);
    let owner = chunk.alloc_scratch(1);
    field(chunk, collection, "owner", line);
    chunk.emit_op_u16(Op::LOCAL_SET, owner, line);
    let axis = chunk.alloc_scratch(1);
    field(chunk, collection, "axis", line);
    chunk.emit_op_u16(Op::LOCAL_SET, axis, line);

    let track = chunk.alloc_scratch(1);
    field(chunk, style, "sizetype", line);
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("auto", line);
    chunk.emit_else(line);
    field(chunk, style, "size", line);
    strings::emit_to_string(chunk, line);
    field(chunk, style, "sizetype", line);
    chunk.emit_call(number, 1, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("px", line);
    chunk.emit_else(line);
    chunk.emit_string_const("fr", line);
    chunk.emit_end(line);
    strings::emit_concat(chunk, 2, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, track, line);

    let old = chunk.alloc_scratch(1);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, owner, line);
    axis_name(chunk, axis, "data-column-styles", "data-row-styles", line);
    let get_attribute = chunk.add_import("web:dom", "getAttribute");
    chunk.emit_call(get_attribute, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, old, line);
    let tracks = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, old, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, old, line);
    chunk.emit_string_const(" ", line);
    chunk.emit_op_u16(Op::LOCAL_GET, track, line);
    strings::emit_concat(chunk, 3, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, track, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, tracks, line);

    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, owner, line);
    axis_name(chunk, axis, "data-column-styles", "data-row-styles", line);
    chunk.emit_op_u16(Op::LOCAL_GET, tracks, line);
    let set_attribute = chunk.add_import("web:dom", "setAttribute");
    chunk.emit_call(set_attribute, 4, line);
    chunk.emit_op(Op::DROP, line);

    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, owner, line);
    axis_name(chunk, axis, "grid-template-columns", "grid-template-rows", line);
    chunk.emit_op_u16(Op::LOCAL_GET, tracks, line);
    let set_style = chunk.add_import("web:cssom", "setStyleProperty");
    chunk.emit_call(set_style, 4, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, collection, line);
    emit_styles_count(chunk, line);
    let index = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, owner, line);
    axis_name(chunk, axis, "data-column-style-count", "data-row-style-count", line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    strings::emit_to_string(chunk, line);
    chunk.emit_call(set_attribute, 4, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
}

/// Stack: `[panel] -> [result]`.
pub fn emit_clear_controls(chunk: &mut Chunk, line: u32) {
    let panel = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, panel, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, panel, line);
    chunk.emit_string_const("", line);
    let clear = chunk.add_import("web:dom", "setInnerHtml");
    chunk.emit_call(clear, 3, line);
}
