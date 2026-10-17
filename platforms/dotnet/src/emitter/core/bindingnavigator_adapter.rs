//! BindingNavigator DOM controls backed by a BindingSource cursor.

use vybe_compiler::primitives::class_slots::{Dest, ObjSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::{class_slots, globals, loops, ops, strings};
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
    // Missing sidecar globals read Undefined, a non-null Wasm externref.
    chunk.emit_op_u16(Op::LOCAL_GET, chunk.dup_slot.expect("dup reserved its root"), line);
    let undefined = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undefined, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_params(line, 1, 1);
    chunk.emit_op(Op::DROP, line);
    let new = chunk.add_import("ecma:object", "new");
    chunk.emit_call(new, 0, line);
    core_wasm::dup(chunk, line);
    globals::emit_write(chunk, SOURCES, line);
    chunk.emit_end(line);
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

/// Keep navigators in sync when their source is populated or moved after binding.
pub fn refresh_source(chunks: &mut [Chunk], current: usize, source: u16, line: u32) {
    source_registry(&mut chunks[current], line);
    let registry = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, registry, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, registry, line);
    let keys_fn = chunks[current].add_import("ecma:object", "keys");
    chunks[current].emit_call(keys_fn, 1, line);
    let keys = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    let count = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);
    let index = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, keys, line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op(Op::ARRAY_GET, line);
    let id = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, id, line);
    c.emit_op_u16(Op::LOCAL_GET, registry, line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    let get = c.add_import("ecma:object", "get");
    c.emit_call(get, 2, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    ops::emit_dyn_eq(c, line);
    c.emit_if(line);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    let get_element = c.add_import("web:dom", "getElementById");
    c.emit_call(get_element, 2, line);
    let navigator = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, navigator, line);
    c.emit_op_u16(Op::LOCAL_GET, navigator, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    refresh(chunks, current, source, navigator, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
}

// Keep the old generated callback display name; the unforgeable metadata
// marker + arity distinguish the private body from all source/wrapper chunks.
const CLICK_BODY_NAME: &str = "__dotnet_bindingnavigator_click";
const CLICK_BODY_MARKER: &str = "<dotnet:bindingnavigator-click-body>";

enum CallbackMove { Static(Move), PrivateLocal(u16) }

fn emit_callback_body(chunks: &mut [Chunk], current: usize, mode: CallbackMove, line: u32) {
    let button = node_from_event(&mut chunks[current], line);
    document(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, button, line);
    let parent = chunks[current].add_import("web:dom", "parentNode");
    chunks[current].emit_call(parent, 2, line);
    let navigator = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, navigator, line);
    let navigator_id = attribute_get(&mut chunks[current], navigator, "id", line);
    globals::emit_read(&mut chunks[current], SOURCES, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, navigator_id, line);
    let get = chunks[current].add_import("ecma:object", "get");
    chunks[current].emit_call(get, 2, line);
    let source = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    let truthy = chunks[current].add_import("ecma:boolean", "toBoolean");
    chunks[current].emit_call(truthy, 1, line);
    let checked_bool = chunks[current].add_import("wasm:js-boolean", "cast");
    chunks[current].emit_call(checked_bool, 1, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    match mode {
        CallbackMove::Static(mode) => bindingsource_adapter::emit_bindingsource_move(chunks, current, mode, line),
        CallbackMove::PrivateLocal(slot) => bindingsource_adapter::emit_bindingsource_move_from_private_mode(chunks, current, slot, line),
    }
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    chunks[current].emit_op(Op::RETURN, line);
}

fn ensure_click_body(chunks: &mut Vec<Chunk>, line: u32) -> Option<usize> {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    if let Some(index) = chunks.iter().position(|c| {
        c.name == CLICK_BODY_NAME && c.arity == 2 && c.param_count == 2
            // Callback errors must retain each installation's original line.
            // Share four modes at this line, never a body from another line.
            && c.get_line(0) == Some(line)
            && c.local_names.iter().any(|n| n.slot == 0
                && n.origin == LocalOrigin::Compiler && n.name == CLICK_BODY_MARKER)
    }) {
        return u16::try_from(index).ok().map(usize::from);
    }
    // Reserve all FOUR new callback indices too. Adding a helper at the last
    // representable slot must not turn formerly representable callbacks wide.
    chunks.len().checked_add(4).filter(|last| *last <= u16::MAX as usize)?;
    let index = chunks.len();
    let mut helper = vybe_compiler::primitives::functions::create_function_chunk(CLICK_BODY_NAME, 2);
    helper.param_count = 2;
    helper.local_count = 2;
    helper.local_names.push(LocalName::new(0, CLICK_BODY_MARKER, LocalOrigin::Compiler));
    let mut body = vec![helper];
    emit_callback_body(&mut body, 0, CallbackMove::PrivateLocal(1), line);
    chunks.push(body.remove(0));
    Some(index)
}

fn callback(chunks: &mut Vec<Chunk>, mode: Move, helper: Option<usize>, line: u32) -> usize {
    let mut function = vybe_compiler::primitives::functions::create_function_chunk(
        "__dotnet_bindingnavigator_click", 1,
    );
    function.local_count = 1;
    if let Some(helper) = helper {
        vybe_compiler::primitives::functions::emit_ref_func(&mut function, helper, 0, line);
        function.emit_op_u16(Op::LOCAL_GET, 0, line);
        function.emit_i32_const(match mode { Move::First=>0, Move::Next=>1, Move::Previous=>2, Move::Last=>3 }, line);
        // Standard tail call replaces this wrapper frame, preserving one
        // generated callback frame and event-root lifetime during refresh.
        function.emit_op_u8_u8(Op::RETURN_CALL_REF, 2, 1, line);
    } else {
        let mut body = vec![function];
        emit_callback_body(&mut body, 0, CallbackMove::Static(mode), line);
        function = body.remove(0);
    }
    chunks.push(function);
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
    let helper = ensure_click_body(chunks, line);
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
        let handler = callback(chunks, mode, helper, line);
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
