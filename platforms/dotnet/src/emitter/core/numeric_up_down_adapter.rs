//! NumericUpDown properties backed by a native number input.

use vybe_compiler::primitives::gui::{DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT};
use vybe_compiler::primitives::strings;
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

pub fn emit_set_value(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);

    let active = chunk.add_import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(active, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    strings::emit_to_string(chunk, line);
    let set = chunk.add_import("web:html", "setValue");
    chunk.emit_call(set, 3, line);
}

pub fn emit_get_value(chunk: &mut Chunk, line: u32) {
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    let active = chunk.add_import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(active, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    let get = chunk.add_import("web:html", "value");
    chunk.emit_call(get, 2, line);
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
}

pub fn emit_get_bound(chunk: &mut Chunk, attribute: &str, line: u32) {
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    let active = chunk.add_import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(active, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const(attribute, line);
    let get = chunk.add_import("web:dom", "getAttribute");
    chunk.emit_call(get, 3, line);
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
}
