//! CheckedListBox items are label rows with independent native checkboxes.

use vybe_compiler::primitives::strings;
use vybe_runtime::{Chunk, opcode::Op, opcode::heaptype::HT_EXTERN};

fn document(chunk: &mut Chunk, line: u32) {
    let active = chunk.add_import("web:html", "activeDocument");
    chunk.emit_call(active, 0, line);
}

fn dom(chunk: &mut Chunk, name: &str, argc: u8, line: u32) {
    let call = chunk.add_import("web:dom", name);
    chunk.emit_call(call, argc, line);
}

fn create(chunk: &mut Chunk, tag: &str, line: u32) {
    document(chunk, line);
    chunk.emit_string_const(tag, line);
    chunk.emit_string_const("", line);
    dom(chunk, "createElement", 3, line);
}

fn append(chunk: &mut Chunk, parent: u16, child: u16, line: u32) {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, parent, line);
    chunk.emit_op_u16(Op::LOCAL_GET, child, line);
    dom(chunk, "appendChild", 3, line);
    chunk.emit_op(Op::DROP, line);
}

fn children(chunk: &mut Chunk, list: u16, line: u32) {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, list, line);
    dom(chunk, "children", 2, line);
}

fn item_checkbox(chunk: &mut Chunk, list: u16, index: u16, line: u32) {
    children(chunk, list, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let row = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, row, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
    dom(chunk, "firstChild", 2, line);
}

/// Stack: `[list, item]` or `[list, item, checked]` -> `[new index]`.
pub fn emit_add(chunk: &mut Chunk, checked_arg: bool, line: u32) {
    let checked = if checked_arg {
        let slot = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        Some(slot)
    } else {
        None
    };
    let item = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, item, line);
    let list = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, list, line);
    children(chunk, list, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    let index = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);

    create(chunk, "label", line);
    let row = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, row, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
    chunk.emit_string_const("style", line);
    chunk.emit_string_const(
        "display:flex;align-items:center;gap:4px;white-space:nowrap;cursor:default",
        line,
    );
    dom(chunk, "setAttribute", 4, line);
    chunk.emit_op(Op::DROP, line);

    create(chunk, "input", line);
    let input = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, input, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_string_const("type", line);
    chunk.emit_string_const("checkbox", line);
    dom(chunk, "setAttribute", 4, line);
    chunk.emit_op(Op::DROP, line);
    if let Some(checked) = checked {
        document(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, input, line);
        chunk.emit_op_u16(Op::LOCAL_GET, checked, line);
        let set = chunk.add_import("web:html", "setChecked");
        chunk.emit_call(set, 3, line);
    }
    append(chunk, row, input, line);

    create(chunk, "span", line);
    let caption = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, caption, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, caption, line);
    chunk.emit_op_u16(Op::LOCAL_GET, item, line);
    strings::emit_to_string(chunk, line);
    dom(chunk, "setTextContent", 3, line);
    chunk.emit_op(Op::DROP, line);
    append(chunk, row, caption, line);
    append(chunk, list, row, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
}

pub fn emit_count(chunk: &mut Chunk, line: u32) {
    let list = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, list, line);
    children(chunk, list, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
}

pub fn emit_clear(chunk: &mut Chunk, line: u32) {
    let list = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, list, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, list, line);
    chunk.emit_string_const("", line);
    dom(chunk, "setTextContent", 3, line);
}

pub fn emit_remove(chunk: &mut Chunk, line: u32) {
    let index = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    let list = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, list, line);
    children(chunk, list, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let row = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, row, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, list, line);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
    dom(chunk, "removeChild", 3, line);
}

pub fn emit_get_checked(chunk: &mut Chunk, line: u32) {
    let index = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    let list = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, list, line);
    item_checkbox(chunk, list, index, line);
    let input = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, input, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    let get = chunk.add_import("web:html", "checked");
    chunk.emit_call(get, 2, line);
}

pub fn emit_set_checked(chunk: &mut Chunk, line: u32) {
    let checked = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, checked, line);
    let index = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    let list = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, list, line);
    item_checkbox(chunk, list, index, line);
    let input = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, input, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_op_u16(Op::LOCAL_GET, checked, line);
    let set = chunk.add_import("web:html", "setChecked");
    chunk.emit_call(set, 3, line);
    chunk.emit_ref_null(HT_EXTERN, line);
}
