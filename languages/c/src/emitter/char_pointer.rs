use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ResolvedSlot};
use vybe_compiler::primitives::instructions::recipes;
use vybe_runtime::{Chunk, Op};

fn local(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn test(chunk: &mut Chunk, module: &str, slot: u16, line: u32) {
    local(chunk, slot, line);
    let test = chunk.add_import(module, "test");
    chunk.emit_call(test, 1, line);
}

fn emit_storage_read(chunk: &mut Chunk, base: u16, index: u16, line: u32) {
    test(chunk, "wasm:js-number", base, line);
    chunk.emit_if_value(line);
    local(chunk, base, line);
    local(chunk, index, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::I32_LOAD8_U, line);
    chunk.emit_else(line);

    test(chunk, "wasm:js-string", base, line);
    chunk.emit_if_value(line);
    local(chunk, index, line);
    local(chunk, base, line);
    let len = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(len, 1, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    local(chunk, base, line);
    local(chunk, index, line);
    let code = chunk.add_import("wasm:js-string", "charCodeAt");
    chunk.emit_call(code, 2, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    local(chunk, index, line);
    // Managed byte storage may be a TypedArray, not a WASM GC array.
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(base),
        &ResolvedSlot::Key("length".into()),
        Dest::Stack,
        line,
    );
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    local(chunk, base, line);
    local(chunk, index, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    test(chunk, "wasm:js-string", value, line);
    chunk.emit_if_value(line);
    local(chunk, value, line);
    chunk.emit_i32_const(0, line);
    let code = chunk.add_import("wasm:js-string", "charCodeAt");
    chunk.emit_call(code, 2, line);
    chunk.emit_else(line);
    local(chunk, value, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub(super) fn emit_read(chunk: &mut Chunk, line: u32) {
    let index = chunk.alloc_scratch(1);
    let pointer = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pointer, line);

    local(chunk, pointer, line);
    recipes::is_object(chunk, line);
    chunk.emit_if_value(line);
    let kind = chunk.alloc_scratch(1);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(pointer),
        &ResolvedSlot::Key("__ref_kind".into()),
        Dest::Local(kind),
        line,
    );
    test(chunk, "wasm:js-string", kind, line);
    chunk.emit_if_value(line);
    local(chunk, kind, line);
    chunk.emit_string_const("carray", line);
    let eq = chunk.add_import("wasm:js-string", "equals");
    chunk.emit_call(eq, 2, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    chunk.emit_if_value(line);

    let base = chunk.alloc_scratch(1);
    let adjusted_index = chunk.alloc_scratch(1);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(pointer),
        &ResolvedSlot::Key("__base".into()),
        Dest::Local(base),
        line,
    );
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(pointer),
        &ResolvedSlot::Key("__idx".into()),
        Dest::Stack,
        line,
    );
    local(chunk, index, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, adjusted_index, line);
    emit_storage_read(chunk, base, adjusted_index, line);
    chunk.emit_else(line);
    emit_storage_read(chunk, pointer, index, line);
    chunk.emit_end(line);
}

pub(super) fn emit_compare(
    chunk: &mut Chunk,
    ascii_case_insensitive: bool,
    nul_terminated: bool,
    bounded: bool,
    line: u32,
) {
    let count = if bounded {
        Some(chunk.alloc_scratch(1))
    } else {
        None
    };
    let right = chunk.alloc_scratch(1);
    let left = chunk.alloc_scratch(1);
    let index = chunk.alloc_scratch(1);
    let left_byte = chunk.alloc_scratch(1);
    let right_byte = chunk.alloc_scratch(1);
    let result = chunk.alloc_scratch(1);

    for slot in count.into_iter().chain([right, left]) {
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);

    let done = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    if let Some(count) = count {
        local(chunk, index, line);
        local(chunk, count, line);
        chunk.emit_op(Op::F64_GE, line);
        chunk.emit_br_if(1, line);
    }

    for (pointer, byte) in [(left, left_byte), (right, right_byte)] {
        local(chunk, pointer, line);
        local(chunk, index, line);
        emit_read(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, byte, line);
        if ascii_case_insensitive {
            local(chunk, byte, line);
            chunk.emit_f64_const(65.0, line);
            chunk.emit_op(Op::F64_GE, line);
            local(chunk, byte, line);
            chunk.emit_f64_const(90.0, line);
            chunk.emit_op(Op::F64_LE, line);
            chunk.emit_op(Op::I32_AND, line);
            chunk.emit_if_value(line);
            local(chunk, byte, line);
            chunk.emit_f64_const(32.0, line);
            chunk.emit_op(Op::F64_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, byte, line);
            chunk.emit_end(line);
        }
    }

    local(chunk, left_byte, line);
    local(chunk, right_byte, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if_value(line);
    local(chunk, left_byte, line);
    local(chunk, right_byte, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_br(2, line);
    chunk.emit_end(line);

    if nul_terminated {
        local(chunk, left_byte, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op(Op::F64_EQ, line);
        chunk.emit_br_if(1, line);
    }

    local(chunk, index, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(done);
    local(chunk, result, line);
}
