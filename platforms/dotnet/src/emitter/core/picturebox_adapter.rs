//! WinForms PictureBox image sources on the canvas-backed control.

use vybe_compiler::primitives::class_slots::{self, ClassSlot, Dest, ObjSource, PlainNames, ResolvedSlot, ValueSource};
use vybe_compiler::primitives::{gui, ops, strings};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

fn stored_property(name: &str) -> ResolvedSlot {
    class_slots::resolve(&ClassSlot::internal(name), &PlainNames)
}

fn set_style(chunk: &mut Chunk, control: u16, name: &str, value: &str, line: u32) {
    let document = chunk.add_import(gui::DOCUMENT_MODULE, gui::HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(document, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const(name, line);
    chunk.emit_string_const(value, line);
    let setter = chunk.add_import("web:cssom", "setStyleProperty");
    chunk.emit_call(setter, 4, line);
    chunk.emit_op(Op::DROP, line);
}

fn store(chunk: &mut Chunk, control: u16, value: u16, name: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(control),
        &stored_property(name),
        ValueSource::Local(value),
        line,
    );
}

pub fn emit_get(chunk: &mut Chunk, name: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &stored_property(name),
        Dest::Stack,
        line,
    );
}

/// Stack: `[pictureBox, location] -> [location]`.
pub fn emit_set_image_location(chunk: &mut Chunk, line: u32) {
    let source = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, source, line);
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    store(chunk, control, source, "__dotnet_image_location", line);

    let document = chunk.add_import(gui::DOCUMENT_MODULE, gui::HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(document, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const("background-image", line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("none", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    strings::emit_to_string(chunk, line);
    chunk.emit_string_const("", line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("none", line);
    chunk.emit_else(line);
    chunk.emit_string_const("url(", line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    let stringify = chunk.add_import("ecma:json", "stringify");
    chunk.emit_call(stringify, 1, line);
    chunk.emit_string_const(")", line);
    strings::emit_concat(chunk, 3, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    let setter = chunk.add_import("web:cssom", "setStyleProperty");
    chunk.emit_call(setter, 4, line);
    chunk.emit_op(Op::DROP, line);
    set_style(chunk, control, "background-repeat", "no-repeat", line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
}

fn is_mode(chunk: &mut Chunk, mode: u16, name: &str, number: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, mode, line);
    strings::emit_to_string(chunk, line);
    chunk.emit_string_const(name, line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, mode, line);
    strings::emit_to_string(chunk, line);
    chunk.emit_string_const(number, line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
}

/// Stack: `[pictureBox, PictureBoxSizeMode] -> [PictureBoxSizeMode]`.
pub fn emit_set_size_mode(chunk: &mut Chunk, line: u32) {
    let mode = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, mode, line);
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    store(chunk, control, mode, "__dotnet_picturebox_size_mode", line);

    is_mode(chunk, mode, "StretchImage", "1", line);
    chunk.emit_if(line);
    set_style(chunk, control, "background-size", "100% 100%", line);
    set_style(chunk, control, "background-position", "center", line);
    chunk.emit_else(line);
    is_mode(chunk, mode, "Zoom", "4", line);
    chunk.emit_if(line);
    set_style(chunk, control, "background-size", "contain", line);
    set_style(chunk, control, "background-position", "center", line);
    chunk.emit_else(line);
    is_mode(chunk, mode, "CenterImage", "3", line);
    chunk.emit_if(line);
    set_style(chunk, control, "background-size", "auto", line);
    set_style(chunk, control, "background-position", "center", line);
    chunk.emit_else(line);
    set_style(chunk, control, "background-size", "auto", line);
    set_style(chunk, control, "background-position", "left top", line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, mode, line);
}
