//! WinForms scroll bars backed by native HTML range inputs.

use vybe_compiler::primitives::strings;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

fn document(c: &mut Chunk, line: u32) {
    let f = c.add_import("web:html", "activeDocument");
    c.emit_call(f, 0, line);
}

fn attribute_name(property: &str) -> Option<&'static str> {
    match property {
        "minimum" => Some("min"),
        "maximum" => Some("max"),
        "smallchange" => Some("step"),
        _ => None,
    }
}

fn set(c: &mut Chunk, control: u16, property: &str, value: u16, line: u32) {
    if property == "largechange" {
        c.emit_op_u16(Op::LOCAL_GET, control, line);
        c.emit_string_const("__dotnet_largechange", line);
        c.emit_op_u16(Op::LOCAL_GET, value, line);
        let f = c.add_import("ecma:object", "set");
        c.emit_call(f, 3, line);
        c.emit_op(Op::DROP, line);
        return;
    }
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, control, line);
    if let Some(attribute) = attribute_name(property) {
        c.emit_string_const(attribute, line);
    }
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    strings::emit_to_string(c, line);
    let (name, argc) = if attribute_name(property).is_some() {
        ("setAttribute", 4)
    } else {
        ("setValue", 3)
    };
    let module = if attribute_name(property).is_some() {
        "web:dom"
    } else {
        "web:html"
    };
    let f = c.add_import(module, name);
    c.emit_call(f, argc, line);
    c.emit_op(Op::DROP, line);
}

pub fn emit_init(c: &mut Chunk, line: u32) {
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    for (property, initial) in [
        ("minimum", 0),
        ("maximum", 100),
        ("smallchange", 1),
        ("largechange", 10),
        ("value", 0),
    ] {
        c.emit_i32_const(initial, line);
        let value = c.alloc_scratch(1);
        c.emit_op_u16(Op::LOCAL_SET, value, line);
        set(c, control, property, value, line);
    }
    c.emit_ref_null(HT_EXTERN, line);
}

pub fn emit_set(c: &mut Chunk, property: &str, line: u32) {
    let value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    set(c, control, property, value, line);
    c.emit_ref_null(HT_EXTERN, line);
}

pub fn emit_get(c: &mut Chunk, property: &str, line: u32) {
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    if property == "largechange" {
        c.emit_op_u16(Op::LOCAL_GET, control, line);
        c.emit_string_const("__dotnet_largechange", line);
        let f = c.add_import("ecma:object", "get");
        c.emit_call(f, 2, line);
        return;
    }
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, control, line);
    let (name, argc) = if let Some(attribute) = attribute_name(property) {
        c.emit_string_const(attribute, line);
        ("getAttribute", 3)
    } else {
        ("value", 2)
    };
    let module = if attribute_name(property).is_some() {
        "web:dom"
    } else {
        "web:html"
    };
    let f = c.add_import(module, name);
    c.emit_call(f, argc, line);
    let number = c.add_import("ecma:number", "Number");
    c.emit_call(number, 1, line);
}
