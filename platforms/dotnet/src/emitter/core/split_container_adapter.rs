//! WinForms SplitContainer panes and splitter on ordinary DOM elements.

use vybe_compiler::primitives::{gui, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

fn document(c: &mut Chunk, line: u32) {
    let f = c.add_import("web:html", "activeDocument");
    c.emit_call(f, 0, line);
}

fn child(c: &mut Chunk, parent: u16, first: bool, line: u32) -> u16 {
    let slot = c.alloc_scratch(1);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, parent, line);
    let f = c.add_import("web:dom", if first { "firstChild" } else { "nextSibling" });
    c.emit_call(f, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn attribute(c: &mut Chunk, node: u16, name: &str, line: u32) -> u16 {
    let slot = c.alloc_scratch(1);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(name, line);
    let f = c.add_import("web:dom", "getAttribute");
    c.emit_call(f, 3, line);
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn set_attribute(c: &mut Chunk, node: u16, name: &str, value: u16, line: u32) {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(name, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let f = c.add_import("web:dom", "setAttribute");
    c.emit_call(f, 4, line);
    c.emit_op(Op::DROP, line);
}

fn style(c: &mut Chunk, node: u16, property: &str, value: u16, line: u32) {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(property, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let f = c.add_import("web:cssom", "setStyleProperty");
    c.emit_call(f, 4, line);
    c.emit_op(Op::DROP, line);
}

fn computed_number(c: &mut Chunk, node: u16, property: &str, line: u32) -> u16 {
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const(property, line);
    let f = c.add_import("web:cssom", "getComputedStyleProperty");
    c.emit_call(f, 3, line);
    let parse = c.add_import("ecma:number", "parseFloat");
    c.emit_call(parse, 1, line);
    let slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn event_node(c: &mut Chunk, key: &str, line: u32) -> u16 {
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_string_const(key, line);
    let get = c.add_import("ecma:object", "get");
    c.emit_call(get, 2, line);
    let new = c.add_import("ecma:object", "new");
    let id = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, id, line);
    c.emit_call(new, 0, line);
    let node = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, node, line);
    c.emit_op_u16(Op::LOCAL_GET, node, line);
    c.emit_string_const("__node", line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    let set = c.add_import("ecma:object", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
    node
}

fn event_coordinate(c: &mut Chunk, line: u32) -> u16 {
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_string_const("clientX", line);
    let get = c.add_import("ecma:object", "get");
    c.emit_call(get, 2, line);
    let number = c.add_import("ecma:number", "Number");
    c.emit_call(number, 1, line);
    let slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn set_distance(c: &mut Chunk, control: u16, distance: u16, line: u32) {
    let pane = child(c, control, true, line);
    c.emit_op_u16(Op::LOCAL_GET, distance, line);
    strings::emit_to_string(c, line);
    let text = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, text, line);
    set_attribute(c, control, "data-splitter-distance", text, line);
    c.emit_string_const("0 0 ", line);
    c.emit_op_u16(Op::LOCAL_GET, text, line);
    ops::emit_dyn_add(c, line);
    c.emit_string_const("px", line);
    ops::emit_dyn_add(c, line);
    let flex = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, flex, line);
    style(c, pane, "flex", flex, line);
}

fn callback(chunks: &mut Vec<Chunk>, name: &str, line: u32) -> usize {
    let mut c = vybe_compiler::primitives::functions::create_function_chunk(name, 1);
    c.local_count = 1;
    let control = event_node(&mut c, "currentTarget", line);
    if name.ends_with("_down") {
        let target = event_node(&mut c, "target", line);
        let target_id = attribute(&mut c, target, "data-splitter", line);
        c.emit_op_u16(Op::LOCAL_GET, target_id, line);
        let truthy = c.add_import("ecma:boolean", "toBoolean");
        c.emit_call(truthy, 1, line);
        c.emit_if(line);
        let x = event_coordinate(&mut c, line);
        c.emit_op_u16(Op::LOCAL_GET, x, line);
        strings::emit_to_string(&mut c, line);
        let x_text = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, x_text, line);
        set_attribute(&mut c, control, "data-splitter-start-x", x_text, line);
        let pane = child(&mut c, control, true, line);
        let width = computed_number(&mut c, pane, "width", line);
        c.emit_op_u16(Op::LOCAL_GET, width, line);
        strings::emit_to_string(&mut c, line);
        let width_text = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, width_text, line);
        set_attribute(&mut c, control, "data-splitter-start-width", width_text, line);
        c.emit_end(line);
    } else if name.ends_with("_move") {
        let start = attribute(&mut c, control, "data-splitter-start-x", line);
        c.emit_op_u16(Op::LOCAL_GET, start, line);
        let truthy = c.add_import("ecma:boolean", "toBoolean");
        c.emit_call(truthy, 1, line);
        c.emit_if(line);
        let x = event_coordinate(&mut c, line);
        c.emit_op_u16(Op::LOCAL_GET, x, line);
        c.emit_op_u16(Op::LOCAL_GET, x, line);
        c.emit_op(Op::F64_EQ, line);
        c.emit_if(line);
        let number = c.add_import("ecma:number", "Number");
        c.emit_op_u16(Op::LOCAL_GET, start, line);
        c.emit_call(number, 1, line);
        c.emit_op_u16(Op::LOCAL_GET, x, line);
        c.emit_op(Op::F64_SUB, line);
        let delta = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, delta, line);
        let initial = attribute(&mut c, control, "data-splitter-start-width", line);
        c.emit_op_u16(Op::LOCAL_GET, initial, line);
        c.emit_call(number, 1, line);
        c.emit_op_u16(Op::LOCAL_GET, delta, line);
        c.emit_op(Op::F64_SUB, line);
        let candidate = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, candidate, line);
        let width = computed_number(&mut c, control, "width", line);
        c.emit_op_u16(Op::LOCAL_GET, width, line);
        c.emit_f64_const(29.0, line);
        c.emit_op(Op::F64_SUB, line);
        let maximum = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, maximum, line);
        c.emit_op_u16(Op::LOCAL_GET, candidate, line);
        c.emit_f64_const(25.0, line);
        c.emit_op(Op::F64_LT, line);
        c.emit_if_value(line);
        c.emit_f64_const(25.0, line);
        c.emit_else(line);
        c.emit_op_u16(Op::LOCAL_GET, candidate, line);
        c.emit_op_u16(Op::LOCAL_GET, maximum, line);
        c.emit_op(Op::F64_GT, line);
        c.emit_if_value(line);
        c.emit_op_u16(Op::LOCAL_GET, maximum, line);
        c.emit_else(line);
        c.emit_op_u16(Op::LOCAL_GET, candidate, line);
        c.emit_end(line);
        c.emit_end(line);
        let distance = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, distance, line);
        set_distance(&mut c, control, distance, line);
        c.emit_end(line);
        c.emit_end(line);
    } else {
        c.emit_string_const("", line);
        let empty = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, empty, line);
        set_attribute(&mut c, control, "data-splitter-start-x", empty, line);
    }
    c.emit_ref_null(HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

/// Stack: [SplitContainer] -> [null].
pub fn emit_init(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    let down = callback(chunks, "__dotnet_split_down", line);
    let moving = callback(chunks, "__dotnet_split_move", line);
    let up = callback(chunks, "__dotnet_split_up", line);
    let c = &mut chunks[current];
    gui::emit_add_event_listener(c, control, "mousedown", down, line);
    gui::emit_add_event_listener(c, control, "mousemove", moving, line);
    gui::emit_add_event_listener(c, control, "mouseup", up, line);
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [SplitContainer, distance] -> [null].
pub fn emit_distance_set(c: &mut Chunk, line: u32) {
    let distance = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, distance, line);
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    set_distance(c, control, distance, line);
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [SplitContainer] -> [distance].
pub fn emit_distance_get(c: &mut Chunk, line: u32) {
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    let pane = child(c, control, true, line);
    let width = computed_number(c, pane, "width", line);
    c.emit_op_u16(Op::LOCAL_GET, width, line);
}

pub fn emit_panel(chunk: &mut Chunk, second: bool, line: u32) {
    let parent = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, parent, line);
    let active = chunk.add_import("web:html", "activeDocument");
    chunk.emit_call(active, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, parent, line);
    let child = chunk.add_import("web:dom", if second { "lastChild" } else { "firstChild" });
    chunk.emit_call(child, 2, line);
}
