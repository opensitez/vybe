//! WinForms BorderStyle enum mapped to CSS without losing the enum on reads.

use vybe_compiler::primitives::gui::{DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT};
use vybe_compiler::primitives::strings;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn document(chunk: &mut Chunk, line: u32) {
    let active = chunk.add_import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(active, 0, line);
}

/// Stack: `[control, BorderStyle] -> [cssom_result]`.
pub fn emit_set_border_style(chunk: &mut Chunk, line: u32) {
    let style = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, style, line);
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);

    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const("data-border-style", line);
    chunk.emit_op_u16(Op::LOCAL_GET, style, line);
    strings::emit_to_string(chunk, line);
    let set_attribute = chunk.add_import("web:dom", "setAttribute");
    chunk.emit_call(set_attribute, 4, line);
    chunk.emit_op(Op::DROP, line);

    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const("border", line);
    chunk.emit_op_u16(Op::LOCAL_GET, style, line);
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("1px solid #7a7a7a", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, style, line);
    chunk.emit_call(number, 1, line);
    chunk.emit_f64_const(2.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("2px inset #d0d0d0", line);
    chunk.emit_else(line);
    chunk.emit_string_const("none", line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    let set_style = chunk.add_import("web:cssom", "setStyleProperty");
    chunk.emit_call(set_style, 4, line);
}

/// Stack: `[control] -> [BorderStyle]`.
pub fn emit_get_border_style(chunk: &mut Chunk, line: u32) {
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const("data-border-style", line);
    let get_attribute = chunk.add_import("web:dom", "getAttribute");
    chunk.emit_call(get_attribute, 3, line);
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
}
