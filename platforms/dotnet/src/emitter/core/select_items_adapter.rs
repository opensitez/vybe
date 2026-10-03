//! WinForms list item collections backed by the select element's options.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

const DOCUMENT: &str = "web:html";

fn document(chunk: &mut Chunk, line: u32) {
    let active = chunk.add_import(DOCUMENT, "activeDocument");
    chunk.emit_call(active, 0, line);
}

/// Stack: `[select] -> [number of options]`.
pub fn emit_count(chunk: &mut Chunk, line: u32) {
    let select = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, select, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, select, line);
    let children = chunk.add_import("web:dom", "children");
    chunk.emit_call(children, 2, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
}

/// Stack: `[select] -> [null]`.
pub fn emit_clear(chunk: &mut Chunk, line: u32) {
    let select = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, select, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, select, line);
    let clear = chunk.add_import("web:html", "clearItems");
    chunk.emit_call(clear, 2, line);
}
