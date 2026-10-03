//! WinForms spacing, docking and text alignment on HTML controls.

use vybe_compiler::primitives::class_slots::{self, ClassSlot, Dest, ObjSource, PlainNames};
use vybe_compiler::primitives::{gui, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn field(chunk: &mut Chunk, object: u16, name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, object, line);
    let slot = class_slots::resolve(&ClassSlot::internal(name), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &slot, Dest::Stack, line);
}

fn set_css(chunk: &mut Chunk, control: u16, property: &str, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    let active = chunk.add_import(gui::DOCUMENT_MODULE, gui::HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(active, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const(property, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let setter = chunk.add_import("web:cssom", "setStyleProperty");
    chunk.emit_call(setter, 4, line);
    chunk.emit_op(Op::DROP, line);
}

/// Stack: `[uniform]` or `[left, top, right, bottom]` -> `[Padding]`.
pub fn emit_padding_new(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc == 1 {
        let value = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
        for _ in 0..4 {
            chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        }
    }
    crate::emitter::dispatch::emit_value_type_new(
        chunk,
        "Padding",
        &["left", "top", "right", "bottom"],
        line,
    );
}

/// Stack: `[control, Padding]` -> `[Padding]`.
pub fn emit_margin_set(chunk: &mut Chunk, line: u32) {
    let padding = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, padding, line);
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    for (index, name) in ["top", "right", "bottom", "left"].iter().enumerate() {
        if index != 0 {
            chunk.emit_string_const(" ", line);
        }
        field(chunk, padding, name, line);
        strings::emit_to_string(chunk, line);
        chunk.emit_string_const("px", line);
        strings::emit_concat(chunk, 2, line);
    }
    strings::emit_concat(chunk, 7, line);
    set_css(chunk, control, "margin", line);
    chunk.emit_op_u16(Op::LOCAL_GET, padding, line);
}

fn enum_is(chunk: &mut Chunk, value: u16, expected: f64, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
    chunk.emit_f64_const(expected, line);
    chunk.emit_op(Op::F64_EQ, line);
}

/// Stack: `[control, DockStyle]` -> `[DockStyle]`.
pub fn emit_dock_set(chunk: &mut Chunk, line: u32) {
    let dock = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, dock, line);
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    enum_is(chunk, dock, 5.0, line);
    chunk.emit_if(line);
    for (property, value) in [
        ("position", "relative"),
        ("width", "100%"),
        ("height", "100%"),
        ("box-sizing", "border-box"),
    ] {
        chunk.emit_string_const(value, line);
        set_css(chunk, control, property, line);
    }
    chunk.emit_else(line);
    enum_is(chunk, dock, 1.0, line);
    chunk.emit_if(line);
    for (property, value) in [("top", "0px"), ("left", "0px"), ("width", "100%")] {
        chunk.emit_string_const(value, line);
        set_css(chunk, control, property, line);
    }
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, dock, line);
}

/// Stack: `[control, ContentAlignment]` -> `[ContentAlignment]`.
pub fn emit_text_align_set(chunk: &mut Chunk, line: u32) {
    let align = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, align, line);
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    chunk.emit_string_const("flex", line);
    set_css(chunk, control, "display", line);
    for expected in [2.0, 32.0, 512.0] {
        enum_is(chunk, align, expected, line);
        chunk.emit_if_value(line);
        chunk.emit_string_const("center", line);
        chunk.emit_else(line);
    }
    chunk.emit_string_const("flex-start", line);
    for _ in 0..3 {
        chunk.emit_end(line);
    }
    set_css(chunk, control, "justify-content", line);
    for expected in [16.0, 32.0, 64.0] {
        enum_is(chunk, align, expected, line);
        chunk.emit_if_value(line);
        chunk.emit_string_const("center", line);
        chunk.emit_else(line);
    }
    chunk.emit_string_const("flex-start", line);
    for _ in 0..3 {
        chunk.emit_end(line);
    }
    set_css(chunk, control, "align-items", line);
    chunk.emit_op_u16(Op::LOCAL_GET, align, line);
}

/// `Tag` is an arbitrary .NET object, not an HTML attribute string.
pub fn emit_tag_set(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_string_const("__dotnet_tag", line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let setter = chunk.add_import("ecma:reflect", "set");
    chunk.emit_call(setter, 3, line);
}

pub fn emit_tag_get(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const("__dotnet_tag", line);
    let getter = chunk.add_import("ecma:reflect", "get");
    chunk.emit_call(getter, 2, line);
}
