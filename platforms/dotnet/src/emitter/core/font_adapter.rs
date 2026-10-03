//! WinForms Font points and style flags mapped to a browser CSS font.

use vybe_compiler::primitives::class_slots::{self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource};
use vybe_compiler::primitives::{gui, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn field(chunk: &mut Chunk, font: u16, name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, font, line);
    let slot = class_slots::resolve(&ClassSlot::internal(name), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &slot, Dest::Stack, line);
}

/// Stack: `[control, Font] -> [cssom_result]`.
pub fn emit_set_control_font(chunk: &mut Chunk, line: u32) {
    let font = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, font, line);
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);

    let font_slot = class_slots::resolve(&ClassSlot::internal("__dotnet_font"), &PlainNames);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(control),
        &font_slot,
        ValueSource::Local(font),
        line,
    );

    let document = chunk.add_import(gui::DOCUMENT_MODULE, gui::HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(document, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const("font", line);

    for (flag, css) in [("italic", "italic "), ("bold", "bold ")] {
        field(chunk, font, flag, line);
        ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        chunk.emit_string_const(css, line);
        chunk.emit_else(line);
        chunk.emit_string_const("", line);
        chunk.emit_end(line);
    }
    field(chunk, font, "size", line);
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
    chunk.emit_f64_const(4.0 / 3.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    strings::emit_to_string(chunk, line);
    chunk.emit_string_const("px ", line);
    field(chunk, font, "name", line);
    strings::emit_to_string(chunk, line);
    chunk.emit_string_const(", Arial, sans-serif", line);
    strings::emit_concat(chunk, 6, line);

    let setter = chunk.add_import("web:cssom", "setStyleProperty");
    chunk.emit_call(setter, 4, line);
}

/// Stack: `[control] -> [Font]`.
pub fn emit_get_control_font(chunk: &mut Chunk, line: u32) {
    let font_slot = class_slots::resolve(&ClassSlot::internal("__dotnet_font"), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &font_slot, Dest::Stack, line);
}

/// Read a public Font property from the canonical drawing value fields.
pub fn emit_get_font_property(chunk: &mut Chunk, property: &str, line: u32) {
    let slot = class_slots::resolve(&ClassSlot::internal(property), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &slot, Dest::Stack, line);
}
