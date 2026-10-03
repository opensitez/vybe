//! WinForms ProgressBar state projected onto HTML's zero-based `<progress>`.

use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::gui::{DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT};
use vybe_compiler::primitives::strings;
use vybe_runtime::{Chunk, opcode::Op, opcode::heaptype::HT_EXTERN};

fn state_slot(property: &str) -> class_slots::ResolvedSlot {
    class_slots::resolve(
        &ClassSlot::internal(&format!("__dotnet_progress_{property}")),
        &PlainNames,
    )
}

fn load_state(chunk: &mut Chunk, control: u16, property: &str, default: i32, line: u32) {
    let slot = state_slot(property);
    class_slots::emit_class_get(chunk, ObjSource::Local(control), &slot, Dest::Stack, line);
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let undefined = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undefined, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(default, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_end(line);
}

fn store_state(chunk: &mut Chunk, control: u16, value: u16, property: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(control),
        &state_slot(property),
        ValueSource::Local(value),
        line,
    );
}

fn set_dom_number(
    chunk: &mut Chunk,
    document: u16,
    control: u16,
    attribute: &str,
    value: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, document, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const(attribute, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    strings::emit_to_string(chunk, line);
    let set = chunk.add_import("web:dom", "setAttribute");
    chunk.emit_call(set, 4, line);
    chunk.emit_op(Op::DROP, line);
}

fn active_document(chunk: &mut Chunk, line: u32) -> u16 {
    let document = chunk.add_import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
    chunk.emit_call(document, 0, line);
    let local = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, local, line);
    local
}

fn remove_value(chunk: &mut Chunk, control: u16, line: u32) {
    let document = active_document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, document, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const("value", line);
    let remove = chunk.add_import("web:dom", "removeAttribute");
    chunk.emit_call(remove, 3, line);
    chunk.emit_op(Op::DROP, line);
}

fn render_value(
    chunk: &mut Chunk,
    document: u16,
    control: u16,
    minimum: u16,
    value: u16,
    line: u32,
) {
    let position = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, minimum, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, position, line);
    set_dom_number(chunk, document, control, "value", position, line);
}

fn render(chunk: &mut Chunk, control: u16, minimum: u16, maximum: u16, value: u16, line: u32) {
    let document_local = active_document(chunk, line);
    let extent = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, maximum, line);
    chunk.emit_op_u16(Op::LOCAL_GET, minimum, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, extent, line);
    // HTML's max must be positive even when the WinForms range has zero width.
    chunk.emit_op_u16(Op::LOCAL_GET, extent, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, extent, line);
    chunk.emit_end(line);
    set_dom_number(chunk, document_local, control, "max", extent, line);

    render_value(chunk, document_local, control, minimum, value, line);
}

fn coerce_i32(chunk: &mut Chunk, line: u32) {
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
}

pub fn emit_get(chunk: &mut Chunk, property: &str, line: u32) {
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    let default = match property {
        "maximum" => 100,
        "step" => 10,
        _ => 0,
    };
    load_state(chunk, control, property, default, line);
}

pub fn emit_set(chunks: &mut [Chunk], current: usize, property: &str, line: u32) {
    let chunk = &mut chunks[current];
    coerce_i32(chunk, line);
    let proposed = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, proposed, line);
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    if property == "step" {
        store_state(chunk, control, proposed, property, line);
        chunk.emit_ref_null(HT_EXTERN, line);
        return;
    }
    if property == "style" {
        for (bound, compare, message) in [
            (0, Op::I32_LT_S, "ProgressBarStyle cannot be negative"),
            (2, Op::I32_GT_S, "ProgressBarStyle exceeds Marquee"),
        ] {
            let chunk = &mut chunks[current];
            chunk.emit_op_u16(Op::LOCAL_GET, proposed, line);
            chunk.emit_i32_const(bound, line);
            chunk.emit_op(compare, line);
            chunk.emit_if(line);
            crate::emitter::core::exceptions::emit_throw_typed(
                chunks,
                current,
                "InvalidEnumArgumentException",
                message,
                line,
            );
            chunks[current].emit_end(line);
        }
        let chunk = &mut chunks[current];
        store_state(chunk, control, proposed, "style", line);
        chunk.emit_op_u16(Op::LOCAL_GET, proposed, line);
        chunk.emit_i32_const(2, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if(line);
        remove_value(chunk, control, line);
        chunk.emit_else(line);
        let minimum = chunk.alloc_scratch(1);
        load_state(chunk, control, "minimum", 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, minimum, line);
        let maximum = chunk.alloc_scratch(1);
        load_state(chunk, control, "maximum", 100, line);
        chunk.emit_op_u16(Op::LOCAL_SET, maximum, line);
        let value = chunk.alloc_scratch(1);
        load_state(chunk, control, "value", 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
        render(chunk, control, minimum, maximum, value, line);
        chunk.emit_end(line);
        chunk.emit_ref_null(HT_EXTERN, line);
        return;
    }

    let minimum = chunk.alloc_scratch(1);
    load_state(chunk, control, "minimum", 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, minimum, line);
    let maximum = chunk.alloc_scratch(1);
    load_state(chunk, control, "maximum", 100, line);
    chunk.emit_op_u16(Op::LOCAL_SET, maximum, line);
    let value = chunk.alloc_scratch(1);
    load_state(chunk, control, "value", 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    if property == "value" {
        for (bound, compare, message) in [
            (minimum, Op::I32_LT_S, "Value is below Minimum"),
            (maximum, Op::I32_GT_S, "Value exceeds Maximum"),
        ] {
            let chunk = &mut chunks[current];
            chunk.emit_op_u16(Op::LOCAL_GET, proposed, line);
            chunk.emit_op_u16(Op::LOCAL_GET, bound, line);
            chunk.emit_op(compare, line);
            chunk.emit_if(line);
            crate::emitter::core::exceptions::emit_throw_typed(
                chunks,
                current,
                "ArgumentException",
                message,
                line,
            );
            chunks[current].emit_end(line);
        }
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, proposed, line);
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    } else {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, proposed, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::I32_LT_S, line);
        chunk.emit_if(line);
        crate::emitter::core::exceptions::emit_throw_typed(
            chunks,
            current,
            "ArgumentException",
            "Progress range cannot be negative",
            line,
        );
        chunks[current].emit_end(line);
        let chunk = &mut chunks[current];
        let (changed, other, compare) = if property == "minimum" {
            (minimum, maximum, Op::I32_GT_S)
        } else {
            (maximum, minimum, Op::I32_LT_S)
        };
        chunk.emit_op_u16(Op::LOCAL_GET, proposed, line);
        chunk.emit_op_u16(Op::LOCAL_SET, changed, line);
        chunk.emit_op_u16(Op::LOCAL_GET, changed, line);
        chunk.emit_op_u16(Op::LOCAL_GET, other, line);
        chunk.emit_op(compare, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, changed, line);
        chunk.emit_op_u16(Op::LOCAL_SET, other, line);
        chunk.emit_end(line);
        for (bound, compare) in [(minimum, Op::I32_LT_S), (maximum, Op::I32_GT_S)] {
            chunk.emit_op_u16(Op::LOCAL_GET, value, line);
            chunk.emit_op_u16(Op::LOCAL_GET, bound, line);
            chunk.emit_op(compare, line);
            chunk.emit_if(line);
            chunk.emit_op_u16(Op::LOCAL_GET, bound, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value, line);
            chunk.emit_end(line);
        }
    }

    let chunk = &mut chunks[current];
    if property == "value" {
        store_state(chunk, control, value, "value", line);
    } else {
        for (name, slot) in [("minimum", minimum), ("maximum", maximum), ("value", value)] {
            store_state(chunk, control, slot, name, line);
        }
    }
    let style = chunk.alloc_scratch(1);
    load_state(chunk, control, "style", 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, style, line);
    chunk.emit_op_u16(Op::LOCAL_GET, style, line);
    chunk.emit_i32_const(2, line);
    chunk.emit_op(Op::I32_NE, line);
    chunk.emit_if(line);
    if property == "value" {
        let document = active_document(chunk, line);
        render_value(chunk, document, control, minimum, value, line);
    } else {
        render(chunk, control, minimum, maximum, value, line);
    }
    chunk.emit_end(line);
    chunk.emit_ref_null(HT_EXTERN, line);
}

/// Stack: `[progress]` or `[progress, delta]` -> `[null]`.
pub fn emit_increment(chunks: &mut [Chunk], current: usize, has_delta: bool, line: u32) {
    let chunk = &mut chunks[current];
    let delta = chunk.alloc_scratch(1);
    if has_delta {
        coerce_i32(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, delta, line);
    }
    let control = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, control, line);
    load_state(chunk, control, "style", 0, line);
    chunk.emit_i32_const(2, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if(line);
    crate::emitter::core::exceptions::emit_throw_typed(
        chunks,
        current,
        "InvalidOperationException",
        "Cannot increment a marquee ProgressBar",
        line,
    );
    chunks[current].emit_end(line);
    let chunk = &mut chunks[current];
    if !has_delta {
        load_state(chunk, control, "step", 10, line);
        chunk.emit_op_u16(Op::LOCAL_SET, delta, line);
    }
    let minimum = chunk.alloc_scratch(1);
    load_state(chunk, control, "minimum", 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, minimum, line);
    let maximum = chunk.alloc_scratch(1);
    load_state(chunk, control, "maximum", 100, line);
    chunk.emit_op_u16(Op::LOCAL_SET, maximum, line);
    let value = chunk.alloc_scratch(1);
    load_state(chunk, control, "value", 0, line);
    chunk.emit_op(Op::I64_EXTEND_I32_S, line);
    chunk.emit_op_u16(Op::LOCAL_GET, delta, line);
    chunk.emit_op(Op::I64_EXTEND_I32_S, line);
    chunk.emit_op(Op::I64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    for (bound, compare) in [(minimum, Op::I64_LT_S), (maximum, Op::I64_GT_S)] {
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        chunk.emit_op_u16(Op::LOCAL_GET, bound, line);
        chunk.emit_op(Op::I64_EXTEND_I32_S, line);
        chunk.emit_op(compare, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, bound, line);
        chunk.emit_op(Op::I64_EXTEND_I32_S, line);
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
        chunk.emit_end(line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    store_state(chunk, control, value, "value", line);
    let document = active_document(chunk, line);
    render_value(chunk, document, control, minimum, value, line);
    chunk.emit_ref_null(HT_EXTERN, line);
}
