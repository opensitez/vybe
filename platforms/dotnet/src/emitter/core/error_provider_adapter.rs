//! WinForms ErrorProvider state and visible validation feedback.

use vybe_compiler::primitives::{collections, loops, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

const ERRORS: &str = "__dotnet_errorprovider_errors";
const TARGETS: &str = "__dotnet_errorprovider_targets";
const OWNER: &str = "__dotnet_errorprovider_owner";
const OUTLINE: &str = "__dotnet_errorprovider_previous_outline";

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

fn nullish(c: &mut Chunk, value: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let f = c.add_import("wasm:js-undefined", "test");
    c.emit_call(f, 1, line);
    c.emit_op(Op::I32_OR, line);
}

fn state(chunks: &mut Vec<Chunk>, current: usize, owner: u16, key: &str, line: u32) -> u16 {
    let c = &mut chunks[current];
    object_get(c, owner, key, line);
    let result = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, result, line);
    nullish(c, result, line);
    c.emit_if(line);
    if key == ERRORS {
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

fn node_key(c: &mut Chunk, target: u16, line: u32) -> u16 {
    object_get(c, target, "__node", line);
    let key = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, key, line);
    key
}

fn dom_attribute(c: &mut Chunk, target: u16, name: &str, value: Option<u16>, line: u32) {
    let doc = c.add_import("web:html", "activeDocument");
    c.emit_call(doc, 0, line);
    c.emit_op_u16(Op::LOCAL_GET, target, line);
    c.emit_string_const(name, line);
    let (operation, argc) = if let Some(value) = value {
        c.emit_op_u16(Op::LOCAL_GET, value, line);
        ("setAttribute", 4)
    } else {
        ("removeAttribute", 3)
    };
    let f = c.add_import("web:dom", operation);
    c.emit_call(f, argc, line);
    c.emit_op(Op::DROP, line);
}

fn style(c: &mut Chunk, target: u16, operation: &str, value: Option<u16>, line: u32) {
    let doc = c.add_import("web:html", "activeDocument");
    c.emit_call(doc, 0, line);
    c.emit_op_u16(Op::LOCAL_GET, target, line);
    c.emit_string_const("outline", line);
    if let Some(value) = value {
        c.emit_op_u16(Op::LOCAL_GET, value, line);
    }
    let f = c.add_import("web:cssom", operation);
    c.emit_call(f, if value.is_some() { 4 } else { 3 }, line);
}

fn clear_visual(c: &mut Chunk, target: u16, line: u32) {
    dom_attribute(c, target, "aria-invalid", None, line);
    dom_attribute(c, target, "data-winforms-error", None, line);
    object_get(c, target, OUTLINE, line);
    let old = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, old, line);
    nullish(c, old, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    style(c, target, "setStyleProperty", Some(old), line);
    c.emit_op(Op::DROP, line);
    c.emit_ref_null(HT_EXTERN, line);
    let empty = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, empty, line);
    object_set(c, target, OUTLINE, empty, line);
    c.emit_end(line);
    c.emit_ref_null(HT_EXTERN, line);
    let empty = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, empty, line);
    object_set(c, target, OWNER, empty, line);
}

/// Stack: [provider, control, message] -> [null].
pub fn emit_set(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    let message = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, message, line);
    let target = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, target, line);
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    let errors = state(chunks, current, owner, ERRORS, line);
    let targets = state(chunks, current, owner, TARGETS, line);
    let c = &mut chunks[current];
    let key = node_key(c, target, line);
    c.emit_op_u16(Op::LOCAL_GET, errors, line);
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
    c.emit_op_u16(Op::LOCAL_GET, errors, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    c.emit_op_u16(Op::LOCAL_GET, message, line);
    let set = c.add_import("ecma:map", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);

    c.emit_op_u16(Op::LOCAL_GET, message, line);
    c.emit_string_const("", line);
    ops::emit_dyn_eq(c, line);
    c.emit_if(line);
    object_get(c, target, OWNER, line);
    c.emit_op_u16(Op::LOCAL_GET, owner, line);
    ops::emit_dyn_eq(c, line);
    c.emit_if(line);
    clear_visual(c, target, line);
    c.emit_end(line);
    c.emit_else(line);
    object_get(c, target, OUTLINE, line);
    let previous = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, previous, line);
    nullish(c, previous, line);
    c.emit_if(line);
    style(c, target, "getStyleProperty", None, line);
    c.emit_op_u16(Op::LOCAL_SET, previous, line);
    object_set(c, target, OUTLINE, previous, line);
    c.emit_end(line);
    object_set(c, target, OWNER, owner, line);
    c.emit_string_const("true", line);
    let yes = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, yes, line);
    dom_attribute(c, target, "aria-invalid", Some(yes), line);
    dom_attribute(c, target, "data-winforms-error", Some(message), line);
    c.emit_string_const("2px solid #c62828", line);
    let outline = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, outline, line);
    style(c, target, "setStyleProperty", Some(outline), line);
    c.emit_op(Op::DROP, line);
    c.emit_end(line);
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [provider, control] -> [message].
pub fn emit_get(c: &mut Chunk, line: u32) {
    let target = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, target, line);
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    let key = node_key(c, target, line);
    object_get(c, owner, ERRORS, line);
    let errors = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, errors, line);
    nullish(c, errors, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, errors, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    let has = c.add_import("ecma:map", "has");
    c.emit_call(has, 2, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, errors, line);
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

/// Stack: [provider] -> [null].
pub fn emit_clear(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    object_get(c, owner, TARGETS, line);
    let targets = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, targets, line);
    nullish(c, targets, line);
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
    clear_visual(c, target, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    chunks[current].emit_end(line);
    let c = &mut chunks[current];
    let new = c.add_import("ecma:map", "new");
    c.emit_call(new, 0, line);
    let errors = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, errors, line);
    object_set(c, owner, ERRORS, errors, line);
    collections::emit_array_new(chunks, current, 0, line);
    let c = &mut chunks[current];
    let targets = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, targets, line);
    object_set(c, owner, TARGETS, targets, line);
    c.emit_ref_null(HT_EXTERN, line);
}
