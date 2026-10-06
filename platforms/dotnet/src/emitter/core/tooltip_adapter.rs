//! WinForms ToolTip state, with the active caption mirrored to HTML `title`.

use vybe_compiler::primitives::{collections, loops, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

const TIPS: &str = "__dotnet_tooltip_values";
const TARGETS: &str = "__dotnet_tooltip_targets";
const OWNER: &str = "__dotnet_tooltip_owner";

fn object_get(c: &mut Chunk, object: u16, key: &str, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, object, line);
    c.emit_string_const(key, line);
    let f = c.add_import("ecma:object", "get");
    c.emit_call(f, 2, line);
}

fn object_set(c: &mut Chunk, object: u16, key: &str, value: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, object, line);
    c.emit_string_const(key, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let f = c.add_import("ecma:object", "set");
    c.emit_call(f, 3, line);
    c.emit_op(Op::DROP, line);
}

fn is_nullish(c: &mut Chunk, value: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let test = c.add_import("wasm:js-undefined", "test");
    c.emit_call(test, 1, line);
    c.emit_op(Op::I32_OR, line);
}

fn node_key(c: &mut Chunk, target: u16, line: u32) -> u16 {
    object_get(c, target, "__node", line);
    let key = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, key, line);
    key
}

fn state(chunks: &mut Vec<Chunk>, current: usize, owner: u16, key: &str, line: u32) -> u16 {
    let c = &mut chunks[current];
    let result = c.alloc_scratch(1);
    object_get(c, owner, key, line);
    c.emit_op_u16(Op::LOCAL_SET, result, line);
    is_nullish(c, result, line);
    c.emit_if(line);
    if key == TIPS {
        let f = c.add_import("ecma:map", "new");
        c.emit_call(f, 0, line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_SET, result, line);
    object_set(c, owner, key, result, line);
    c.emit_end(line);
    result
}

fn title(c: &mut Chunk, target: u16, text: Option<u16>, line: u32) {
    let doc = c.add_import("web:html", "activeDocument");
    c.emit_call(doc, 0, line);
    c.emit_op_u16(Op::LOCAL_GET, target, line);
    c.emit_string_const("title", line);
    let (method, argc) = if let Some(text) = text {
        c.emit_op_u16(Op::LOCAL_GET, text, line);
        ("setAttribute", 4)
    } else {
        ("removeAttribute", 3)
    };
    let f = c.add_import("web:dom", method);
    c.emit_call(f, argc, line);
    c.emit_op(Op::DROP, line);
}

/// The hidden element carries a Component, not a visual Control. Its defaults
/// are typed fields so .NET property reads do not inherit HTML string coercion.
pub fn emit_init(c: &mut Chunk, line: u32) {
    let tooltip = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, tooltip, line);
    for (name, value) in [
        ("active", true),
        ("showalways", false),
        ("usefading", true),
        ("useanimation", true),
    ] {
        c.emit_bool_const(value, line);
        let slot = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
        object_set(c, tooltip, name, slot, line);
    }
    for (name, value) in [
        ("automaticdelay", 500),
        ("autopopdelay", 5000),
        ("initialdelay", 500),
        ("reshowdelay", 100),
    ] {
        c.emit_i32_const(value, line);
        let slot = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
        object_set(c, tooltip, name, slot, line);
    }
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [tooltip, control, caption] -> [null].
pub fn emit_set(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    let caption = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, caption, line);
    let target = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, target, line);
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    let tips = state(chunks, current, owner, TIPS, line);
    let targets = state(chunks, current, owner, TARGETS, line);
    let c = &mut chunks[current];
    let key = node_key(c, target, line);
    c.emit_op_u16(Op::LOCAL_GET, tips, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    let has = c.add_import("ecma:map", "has");
    c.emit_call(has, 2, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, targets, line);
    c.emit_op_u16(Op::LOCAL_GET, target, line);
    let push = c.add_import("ecma:array", "push");
    c.emit_call(push, 2, line);
    c.emit_op(Op::DROP, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, tips, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    c.emit_op_u16(Op::LOCAL_GET, caption, line);
    let set = c.add_import("ecma:map", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
    object_set(c, target, OWNER, owner, line);
    object_get(c, owner, "active", line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    title(c, target, Some(caption), line);
    c.emit_else(line);
    title(c, target, None, line);
    c.emit_end(line);
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [tooltip, Boolean] -> [null]. Stored captions survive deactivation.
pub fn emit_active_set(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    let active = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, active, line);
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    object_set(c, owner, "active", active, line);
    object_get(c, owner, TARGETS, line);
    let targets = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, targets, line);
    is_nullish(c, targets, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    object_get(c, owner, TIPS, line);
    let tips = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, tips, line);
    c.emit_i32_const(0, line);
    let index = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op_u16(Op::LOCAL_GET, targets, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, targets, line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op(Op::ARRAY_GET, line);
    let target = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, target, line);
    c.emit_op_u16(Op::LOCAL_GET, active, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, tips, line);
    let key = node_key(c, target, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    let get = c.add_import("ecma:map", "get");
    c.emit_call(get, 2, line);
    let caption = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, caption, line);
    title(c, target, Some(caption), line);
    c.emit_else(line);
    title(c, target, None, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// Stack: [tooltip, milliseconds] -> [null].
pub fn emit_automatic_delay_set(c: &mut Chunk, line: u32) {
    let delay = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, delay, line);
    let tooltip = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, tooltip, line);
    for name in ["automaticdelay", "initialdelay"] {
        object_set(c, tooltip, name, delay, line);
    }

    c.emit_op_u16(Op::LOCAL_GET, delay, line);
    c.emit_i32_const(5, line);
    c.emit_op(Op::I32_DIV_S, line);
    let reshow = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, reshow, line);
    object_set(c, tooltip, "reshowdelay", reshow, line);

    c.emit_op_u16(Op::LOCAL_GET, delay, line);
    c.emit_i32_const(10, line);
    c.emit_op(Op::I32_MUL, line);
    let autopop = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, autopop, line);
    object_set(c, tooltip, "autopopdelay", autopop, line);
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [tooltip, control] -> [caption].
pub fn emit_get(c: &mut Chunk, line: u32) {
    let target = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, target, line);
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    let key = node_key(c, target, line);
    object_get(c, owner, TIPS, line);
    let tips = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, tips, line);
    is_nullish(c, tips, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, tips, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    let has = c.add_import("ecma:map", "has");
    c.emit_call(has, 2, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, tips, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    let get = c.add_import("ecma:map", "get");
    c.emit_call(get, 2, line);
    c.emit_else(line);
    c.emit_string_const("", line);
    c.emit_end(line);
    c.emit_else(line);
    c.emit_string_const("", line);
    c.emit_end(line);
}

/// Stack: [tooltip] -> [null].
pub fn emit_remove_all(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    object_get(c, owner, TARGETS, line);
    let targets = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, targets, line);
    is_nullish(c, targets, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_i32_const(0, line);
    let index = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op_u16(Op::LOCAL_GET, targets, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, targets, line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op(Op::ARRAY_GET, line);
    let target = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, target, line);
    object_get(c, target, OWNER, line);
    c.emit_op_u16(Op::LOCAL_GET, owner, line);
    ops::emit_dyn_eq(c, line);
    c.emit_if(line);
    title(c, target, None, line);
    c.emit_ref_null(HT_EXTERN, line);
    let empty = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, empty, line);
    object_set(c, target, OWNER, empty, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    chunks[current].emit_end(line);
    let c = &mut chunks[current];
    let map = c.add_import("ecma:map", "new");
    c.emit_call(map, 0, line);
    let tips = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, tips, line);
    object_set(c, owner, TIPS, tips, line);
    collections::emit_array_new(chunks, current, 0, line);
    let c = &mut chunks[current];
    let targets = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, targets, line);
    object_set(c, owner, TARGETS, targets, line);
    c.emit_ref_null(HT_EXTERN, line);
}
