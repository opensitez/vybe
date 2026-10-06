//! Per-control WinForms HelpProvider metadata on nonvisual components.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

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

fn control_key(c: &mut Chunk, control: u16, line: u32) -> u16 {
    object_get(c, control, "__node", line);
    let key = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, key, line);
    key
}

fn map_name(property: &str) -> &'static str {
    match property {
        "string" => "__dotnet_help_strings",
        "keyword" => "__dotnet_help_keywords",
        "navigator" => "__dotnet_help_navigators",
        "show" => "__dotnet_help_visible",
        _ => unreachable!("registered HelpProvider property"),
    }
}

fn default(c: &mut Chunk, property: &str, line: u32) {
    if property == "show" {
        c.emit_bool_const(false, line);
    } else {
        c.emit_string_const("", line);
    }
}

/// Stack: [provider, control, value] -> [null].
pub fn emit_set(c: &mut Chunk, property: &str, line: u32) {
    let value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    let provider = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, provider, line);
    object_get(c, provider, map_name(property), line);
    let map = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, map, line);
    nullish(c, map, line);
    c.emit_if(line);
    let new = c.add_import("ecma:map", "new");
    c.emit_call(new, 0, line);
    c.emit_op_u16(Op::LOCAL_SET, map, line);
    object_set(c, provider, map_name(property), map, line);
    c.emit_end(line);
    let key = control_key(c, control, line);
    c.emit_op_u16(Op::LOCAL_GET, map, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let set = c.add_import("ecma:map", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
    if property == "string" {
        let doc = c.add_import("web:html", "activeDocument");
        c.emit_call(doc, 0, line);
        c.emit_op_u16(Op::LOCAL_GET, control, line);
        c.emit_string_const("aria-description", line);
        c.emit_op_u16(Op::LOCAL_GET, value, line);
        let set = c.add_import("web:dom", "setAttribute");
        c.emit_call(set, 4, line);
        c.emit_op(Op::DROP, line);
    }
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [provider, control] -> [value].
pub fn emit_get(c: &mut Chunk, property: &str, line: u32) {
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    let provider = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, provider, line);
    let key = control_key(c, control, line);
    object_get(c, provider, map_name(property), line);
    let map = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, map, line);
    nullish(c, map, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, map, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    let has = c.add_import("ecma:map", "has");
    c.emit_call(has, 2, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, map, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    let get = c.add_import("ecma:map", "get");
    c.emit_call(get, 2, line);
    c.emit_else(line);
    default(c, property, line);
    c.emit_end(line);
    c.emit_else(line);
    default(c, property, line);
    c.emit_end(line);
}

pub fn emit_init(c: &mut Chunk, line: u32) {
    let provider = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, provider, line);
    c.emit_string_const("", line);
    let value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    object_set(c, provider, "helpnamespace", value, line);
    c.emit_ref_null(HT_EXTERN, line);
}
