//! System.Drawing.Color values and CSS color properties over the browser CSSOM.

use vybe_compiler::primitives::class_slots::{self, ClassSlot, Dest, ObjSource, PlainNames};
use vybe_compiler::primitives::{gui, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const DOM: &str = "web:dom";
const CSSOM: &str = "web:cssom";

fn document(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import(gui::DOCUMENT_MODULE, gui::HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(idx, 0, line);
}

fn call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

fn color_value(chunk: &mut Chunk, line: u32) {
    crate::emitter::dispatch::emit_value_type_new(chunk, "Color", &["r", "g", "b", "a"], line);
}

/// Stack: `[r, g, b]` or `[a, r, g, b]` -> `[Color]`.
pub fn emit_from_argb(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc == 3 {
        chunk.emit_f64_const(255.0, line);
    } else {
        let channels = chunk.alloc_scratch(4);
        for index in (0..4).rev() {
            chunk.emit_op_u16(Op::LOCAL_SET, channels + index, line);
        }
        for index in [1, 2, 3, 0] {
            chunk.emit_op_u16(Op::LOCAL_GET, channels + index, line);
        }
    }
    color_value(chunk, line);
}

fn parse_hex_component(chunk: &mut Chunk, text: u16, start: f64, digits: u8, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    chunk.emit_f64_const(start, line);
    chunk.emit_f64_const(start + digits as f64, line);
    call(chunk, "ecma:string", "substring", 3, line);
    if digits == 1 {
        let digit = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, digit, line);
        chunk.emit_op_u16(Op::LOCAL_GET, digit, line);
        chunk.emit_op_u16(Op::LOCAL_GET, digit, line);
        strings::emit_str_concat(chunk, line);
    }
    chunk.emit_f64_const(16.0, line);
    call(chunk, "ecma:number", "parseInt", 2, line);
}

fn hex_color(chunk: &mut Chunk, text: u16, line: u32) {
    let channels = chunk.alloc_scratch(3);
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    call(chunk, "ecma:string", "length", 1, line);
    chunk.emit_f64_const(4.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);
    for (index, start) in [1.0, 2.0, 3.0].into_iter().enumerate() {
        parse_hex_component(chunk, text, start, 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, channels + index as u16, line);
    }
    chunk.emit_else(line);
    for (index, start) in [1.0, 3.0, 5.0].into_iter().enumerate() {
        parse_hex_component(chunk, text, start, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, channels + index as u16, line);
    }
    chunk.emit_end(line);
    for index in 0..3 {
        chunk.emit_op_u16(Op::LOCAL_GET, channels + index, line);
    }
    chunk.emit_f64_const(255.0, line);
    color_value(chunk, line);
}

fn rgb_component(chunk: &mut Chunk, parts: u16, index: f64, alpha: bool, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, parts, line);
    chunk.emit_f64_const(index, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    call(
        chunk,
        "ecma:number",
        if alpha { "parseFloat" } else { "parseInt" },
        1,
        line,
    );
    if alpha {
        chunk.emit_f64_const(255.0, line);
        chunk.emit_op(Op::F64_MUL, line);
    }
}

fn rgb_color(chunk: &mut Chunk, text: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    chunk.emit_string_const("rgba", line);
    call(chunk, "ecma:string", "startsWith", 2, line);
    let has_alpha = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, has_alpha, line);

    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    chunk.emit_op_u16(Op::LOCAL_GET, has_alpha, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(5.0, line);
    chunk.emit_else(line);
    chunk.emit_f64_const(4.0, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    call(chunk, "ecma:string", "length", 1, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    call(chunk, "ecma:string", "substring", 3, line);
    chunk.emit_string_const(",", line);
    call(chunk, "ecma:string", "split", 2, line);
    let parts = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, parts, line);
    for index in [0.0, 1.0, 2.0] {
        rgb_component(chunk, parts, index, false, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, has_alpha, line);
    chunk.emit_if_value(line);
    rgb_component(chunk, parts, 3.0, true, line);
    chunk.emit_else(line);
    chunk.emit_f64_const(255.0, line);
    chunk.emit_end(line);
    color_value(chunk, line);
}

fn browser_color(chunk: &mut Chunk, source: u16, line: u32) {
    let node = chunk.alloc_scratch(1);
    document(chunk, line);
    chunk.emit_string_const("span", line);
    chunk.emit_string_const("", line);
    call(chunk, DOM, "createElement", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, node, line);
    for (property, value) in [("display", None), ("color", Some(source))] {
        document(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, node, line);
        chunk.emit_string_const(property, line);
        if let Some(slot) = value {
            chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        } else {
            chunk.emit_string_const("none", line);
        }
        call(chunk, CSSOM, "setStyleProperty", 4, line);
        chunk.emit_op(Op::DROP, line);
    }
    document(chunk, line);
    document(chunk, line);
    call(chunk, gui::DOCUMENT_MODULE, "body", 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    call(chunk, DOM, "appendChild", 3, line);
    chunk.emit_op(Op::DROP, line);

    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    chunk.emit_string_const("color", line);
    call(chunk, CSSOM, "getComputedStyleProperty", 3, line);
    let computed = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, computed, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, node, line);
    call(chunk, DOM, "remove", 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, computed, line);
    chunk.emit_string_const("#", line);
    call(chunk, "ecma:string", "startsWith", 2, line);
    chunk.emit_if_value(line);
    hex_color(chunk, computed, line);
    chunk.emit_else(line);
    rgb_color(chunk, computed, line);
    chunk.emit_end(line);
}

/// Stack: `[html_color] -> [System.Drawing.Color]`.
pub fn emit_from_html(chunk: &mut Chunk, line: u32) {
    strings::emit_to_string(chunk, line);
    let source = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, source, line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    chunk.emit_string_const("#", line);
    call(chunk, "ecma:string", "startsWith", 2, line);
    chunk.emit_if_value(line);
    hex_color(chunk, source, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    chunk.emit_string_const("rgb", line);
    call(chunk, "ecma:string", "startsWith", 2, line);
    chunk.emit_if_value(line);
    rgb_color(chunk, source, line);
    chunk.emit_else(line);
    browser_color(chunk, source, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn from_html_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(index) = chunks
        .iter()
        .position(|chunk| chunk.name == "__dotnet_color_from_html")
    {
        return index;
    }
    let mut helper = Chunk::new("__dotnet_color_from_html");
    helper.arity = 1;
    helper.emit_op_u16(Op::LOCAL_GET, 0, line);
    emit_from_html(&mut helper, line);
    helper.emit_op(Op::RETURN, line);
    chunks.push(helper);
    chunks.len() - 1
}

/// Stack: `[html_color] -> [System.Drawing.Color]`.
pub fn emit_from_html_call(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let index = from_html_chunk(chunks, line);
    let chunk = &mut chunks[current];
    let source = chunk.alloc_scratch(2);
    let function = source + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, source, line);
    chunk.emit_op_u16(Op::REF_FUNC, index as u16, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, function, line);
    chunk.emit_op_u16(Op::LOCAL_GET, function, line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
}

/// Stack: `[element] -> [Color]`.
pub fn emit_get_control_color(
    chunks: &mut Vec<Chunk>,
    current: usize,
    background: bool,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const(
        if background {
            "background-color"
        } else {
            "color"
        },
        line,
    );
    call(chunk, CSSOM, "getStyleProperty", 3, line);
    emit_from_html_call(chunks, current, line);
}

fn channel(chunk: &mut Chunk, color: u16, field: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, color, line);
    let slot = class_slots::resolve(&ClassSlot::internal(field), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &slot, Dest::Stack, line);
}

/// Stack: `[element, Color] -> [cssom_result]`.
pub fn emit_set_control_color(chunk: &mut Chunk, background: bool, line: u32) {
    let color = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, color, line);
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const(
        if background {
            "background-color"
        } else {
            "color"
        },
        line,
    );
    chunk.emit_string_const("rgba(", line);
    for field in ["r", "g", "b"] {
        channel(chunk, color, field, line);
        strings::emit_to_string(chunk, line);
        chunk.emit_string_const(",", line);
    }
    channel(chunk, color, "a", line);
    chunk.emit_f64_const(255.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    strings::emit_to_string(chunk, line);
    chunk.emit_string_const(")", line);
    strings::emit_concat(chunk, 9, line);
    call(chunk, CSSOM, "setStyleProperty", 4, line);
}
