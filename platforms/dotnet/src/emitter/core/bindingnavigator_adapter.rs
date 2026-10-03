//! BindingNavigator DOM controls backed by a BindingSource cursor.

use vybe_compiler::primitives::class_slots::{Dest, ObjSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::{class_slots, globals, ops, strings};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, opcode::heaptype::HT_EXTERN};

use super::bindingsource_adapter::{self, Move};
use super::object_fields::field_slot;

const SOURCES: &str = "__dotnet_bindingnavigator_sources";

fn document(chunk: &mut Chunk, line: u32) {
    let active = chunk.add_import("web:html", "activeDocument");
    chunk.emit_call(active, 0, line);
}

fn child(chunk: &mut Chunk, node: u16, first: bool, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    let method = chunk.add_import("web:dom", if first { "firstChild" } else { "nextSibling" });
    chunk.emit_call(method, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn attribute_get(chunk: &mut Chunk, node: u16, name: &str, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const(name, line);
    let get = chunk.add_import("web:dom", "getAttribute");
    chunk.emit_call(get, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn source_registry(chunk: &mut Chunk, line: u32) {
    globals::emit_read(chunk, SOURCES, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let create_block = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);
    let new = chunk.add_import("ecma:object", "new");
    chunk.emit_call(new, 0, line);
    core_wasm::dup(chunk, line);
    globals::emit_write(chunk, SOURCES, line);
    chunk.emit_end(line);
    chunk.patch_block(create_block);
}

fn node_from_event(chunk: &mut Chunk, line: u32) -> u16 {
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    chunk.emit_string_const("currentTarget", line);
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_call(get, 2, line);
    let node_id = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, node_id, line);
    let new = chunk.add_import("ecma:object", "new");
    chunk.emit_call(new, 0, line);
    let node = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, node, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const("__node", line);
    chunk.emit_op_u16(Op::LOCAL_GET, node_id, line);
    let set = chunk.add_import("ecma:object", "set");
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
    node
}

fn refresh(chunks: &mut [Chunk], current: usize, source: u16, navigator: u16, line: u32) {
    let count = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    bindingsource_adapter::emit_bindingsource_count(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);

    let position = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Stack,
        &field_slot("position"),
        Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, position, line);

    let input = {
        let c = &mut chunks[current];
        let first = child(c, navigator, true, line);
        let second = child(c, first, false, line);
        child(c, second, false, line)
    };
    let c = &mut chunks[current];
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, input, line);
    c.emit_op_u16(Op::LOCAL_GET, count, line);
    c.emit_i32_const(0, line);
    ops::emit_dyn_eq(c, line);
    c.emit_if_value(line);
    c.emit_string_const("0", line);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, position, line);
    c.emit_f64_const(1.0, line);
    ops::emit_dyn_add(c, line);
    strings::emit_to_string(c, line);
    c.emit_end(line);
    c.emit_string_const(" of ", line);
    ops::emit_dyn_add(c, line);
    c.emit_op_u16(Op::LOCAL_GET, count, line);
    strings::emit_to_string(c, line);
    ops::emit_dyn_add(c, line);
    let set_value = c.add_import("web:html", "setValue");
    c.emit_call(set_value, 3, line);
    c.emit_op(Op::DROP, line);
}

fn callback(chunks: &mut Vec<Chunk>, mode: Move, line: u32) -> usize {
    let mut function = vybe_compiler::primitives::functions::create_function_chunk(
        "__dotnet_bindingnavigator_click",
        1,
    );
    function.local_count = 1;
    let mut body = vec![function];
    let button = node_from_event(&mut body[0], line);
    document(&mut body[0], line);
    body[0].emit_op_u16(Op::LOCAL_GET, button, line);
    let parent = body[0].add_import("web:dom", "parentNode");
    body[0].emit_call(parent, 2, line);
    let navigator = body[0].alloc_scratch(1);
    body[0].emit_op_u16(Op::LOCAL_SET, navigator, line);
    let navigator_id = attribute_get(&mut body[0], navigator, "id", line);
    globals::emit_read(&mut body[0], SOURCES, line);
    body[0].emit_op_u16(Op::LOCAL_GET, navigator_id, line);
    let get = body[0].add_import("ecma:object", "get");
    body[0].emit_call(get, 2, line);
    let source = body[0].alloc_scratch(1);
    body[0].emit_op_u16(Op::LOCAL_SET, source, line);
    body[0].emit_op_u16(Op::LOCAL_GET, source, line);
    let truthy = body[0].add_import("ecma:boolean", "toBoolean");
    body[0].emit_call(truthy, 1, line);
    body[0].emit_if(line);
    body[0].emit_op_u16(Op::LOCAL_GET, source, line);
    bindingsource_adapter::emit_bindingsource_move(&mut body, 0, mode, line);
    body[0].emit_op(Op::DROP, line);
    refresh(&mut body, 0, source, navigator, line);
    body[0].emit_end(line);
    body[0].emit_ref_null(HT_EXTERN, line);
    body[0].emit_op(Op::RETURN, line);
    chunks.push(body.remove(0));
    chunks.len() - 1
}

/// Stack: [navigator, source] -> [null].
pub fn emit_set_source(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source, line);
    let navigator = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, navigator, line);

    let c = &mut chunks[current];
    source_registry(c, line);
    let registry = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, registry, line);
    let navigator_id = attribute_get(c, navigator, "id", line);
    c.emit_op_u16(Op::LOCAL_GET, registry, line);
    c.emit_op_u16(Op::LOCAL_GET, navigator_id, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    let set = c.add_import("ecma:object", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
    let mut button = child(&mut chunks[current], navigator, true, line);
    for (index, mode) in [
        (0, Move::First),
        (1, Move::Previous),
        (3, Move::Next),
        (4, Move::Last),
    ] {
        if index != 0 {
            button = child(&mut chunks[current], button, false, line);
            if index == 3 {
                button = child(&mut chunks[current], button, false, line);
            }
        }
        let handler = callback(chunks, mode, line);
        let c = &mut chunks[current];
        vybe_compiler::primitives::gui::emit_add_event_listener(c, button, "click", handler, line);
    }
    refresh(chunks, current, source, navigator, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

pub fn emit_get_source(chunk: &mut Chunk, line: u32) {
    let navigator = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, navigator, line);
    let navigator_id = attribute_get(chunk, navigator, "id", line);
    source_registry(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, navigator_id, line);
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_call(get, 2, line);
}
