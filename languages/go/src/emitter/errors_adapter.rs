use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::{callable, class_slots, collections, expressions, ops, reflection};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

const GO_ERROR_FLAG: &str = "__go_error";
const GO_ERROR_TYPE: &str = "__goError";

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.errors_new" => {
            if argc != 1 {
                return false;
            }
            chunks[current].emit_ref_null(HT_EXTERN, line);
            chunks[current].emit_ref_null(HT_EXTERN, line);
            emit_go_error_new_from_stack(chunks, current, line);
        }
        "go.errors_errorf" => {
            if argc != 3 {
                return false;
            }
            emit_go_error_new_from_stack(chunks, current, line);
        }
        "go.errors_string" => {
            if argc != 1 {
                return false;
            }
            emit_go_error_string_from_stack(&mut chunks[current], line);
        }
        "go.errors_unwrap" => {
            if argc != 1 {
                return false;
            }
            emit_go_errors_unwrap(chunks, current, line);
        }
        "go.errors_is" => {
            if argc != 2 {
                return false;
            }
            emit_go_errors_is(chunks, current, line);
        }
        "go.errors_join" => {
            if argc == 1 {
                emit_go_errors_join(chunks, current, line);
            } else {
                let base = chunks[current].local_count;
                chunks[current].alloc_scratch(argc as u16);
                collections::emit_pack_n(chunks, current, argc as u16, base, line);
                emit_go_errors_join(chunks, current, line);
            }
        }
        "go.errors_as" => {
            if argc != 3 {
                return false;
            }
            emit_go_errors_as(chunks, current, line);
        }
        _ => return false,
    }
    true
}

pub fn emit_go_error_string_value(chunk: &mut Chunk, value_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("", line);
    chunk.emit_else(line);
    emit_is_go_error_slot(chunk, value_slot, line);
    chunk.emit_if_value(line);
    emit_get_field_slot(chunk, value_slot, "message", line);
    chunk.emit_else(line);
    expressions::emit_rich_to_string(chunk, value_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_go_error_new_from_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let obj = base;
    let msg = base + 1;
    let wrap = base + 2;
    let errs = base + 3;
    let types = base + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, errs, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, wrap, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, msg, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, msg, line);
    let string_ctor = chunks[current].add_import("ecma:string", "String");
    chunks[current].emit_call(string_ctor, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, msg, line);

    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj, line);

    set_field_local(&mut chunks[current], obj, "message", msg, line);
    set_field_local(&mut chunks[current], obj, "wrap", wrap, line);
    set_field_local(&mut chunks[current], obj, "errs", errs, line);
    set_field_bool(&mut chunks[current], obj, GO_ERROR_FLAG, true, line);
    set_field_str(
        &mut chunks[current],
        obj,
        reflection::FIELD_TYPE,
        GO_ERROR_TYPE,
        line,
    );
    set_field_str(
        &mut chunks[current],
        obj,
        reflection::FIELD_TYPE_NAME,
        GO_ERROR_TYPE,
        line,
    );
    set_field_str(
        &mut chunks[current],
        obj,
        reflection::FIELD_KIND,
        "error",
        line,
    );

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, types, line);
    for ty in [GO_ERROR_TYPE, "error", "Exception"] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, types, line);
        chunks[current].emit_string_const(ty, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    set_field_local(
        &mut chunks[current],
        obj,
        reflection::FIELD_TYPES,
        types,
        line,
    );

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
}

fn emit_go_error_string_from_stack(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    emit_go_error_string_value(chunk, value, line);
}

fn emit_go_errors_unwrap(chunks: &mut [Chunk], current: usize, line: u32) {
    let err = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, err, line);
    emit_is_go_error_slot(&mut chunks[current], err, line);
    chunks[current].emit_if_value(line);
    emit_get_field_slot(&mut chunks[current], err, "wrap", line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    chunks[current].emit_end(line);
}

fn emit_go_errors_is(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let err = base;
    let target = base + 1;
    let work = base + 2;
    let cur = base + 3;
    let found = base + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, target, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, err, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, work, line);
    emit_push_local_if_non_null(chunks, current, work, err, line);
    emit_error_chain_scan(chunks, current, work, cur, found, None, None, line);
    emit_bool_from_i32_slot(&mut chunks[current], found, line);
    chunks[current].emit_end(line);
}

fn emit_go_errors_join(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(7);
    let input = base;
    let filtered = base + 1;
    let msg = base + 2;
    let i = base + 3;
    let len = base + 4;
    let cur = base + 5;
    let count = base + 6;

    chunks[current].emit_op_u16(Op::LOCAL_SET, input, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, filtered, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, msg, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);
    emit_array_len_slot(&mut chunks[current], input, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    let outer = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    emit_array_get_slot(chunks, current, input, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, msg, line);
    chunks[current].emit_string_const("\n", line);
    host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, msg, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, msg, line);
    emit_go_error_string_value(&mut chunks[current], cur, line);
    host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, msg, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, filtered, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, msg, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, filtered, line);
    emit_go_error_new_from_stack(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_go_errors_as(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    let err = base;
    let matcher = base + 1;
    let assigner = base + 2;
    let work = base + 3;
    let cur = base + 4;
    let found = base + 5;

    chunks[current].emit_op_u16(Op::LOCAL_SET, assigner, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, matcher, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, err, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, work, line);
    emit_push_local_if_non_null(chunks, current, work, err, line);
    emit_error_chain_scan(
        chunks,
        current,
        work,
        cur,
        found,
        Some(matcher),
        Some(assigner),
        line,
    );
    emit_bool_from_i32_slot(&mut chunks[current], found, line);
}

fn emit_error_chain_scan(
    chunks: &mut [Chunk],
    current: usize,
    work: u16,
    cur: u16,
    found: u16,
    matcher: Option<u16>,
    assigner: Option<u16>,
    line: u32,
) {
    let outer = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    emit_array_len_slot(&mut chunks[current], work, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, work, line);
    let pop = chunks[current].add_import("ecma:array", "pop");
    chunks[current].emit_call(pop, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    if let Some(matcher) = matcher {
        callable::emit_callback_on(chunks, current, matcher, cur, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    } else {
        let target = work - 1;
        chunks[current].emit_op_u16(Op::LOCAL_GET, cur, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
        ops::emit_dyn_eq(&mut chunks[current], line);
    }
    chunks[current].emit_if(line);
    if let (Some(_), Some(assigner)) = (matcher, assigner) {
        callable::emit_callback_on(chunks, current, assigner, cur, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);
    chunks[current].emit_else(line);
    emit_push_go_error_children(chunks, current, work, cur, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
}

fn emit_push_go_error_children(
    chunks: &mut [Chunk],
    current: usize,
    work: u16,
    cur: u16,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(3);
    let errs = base;
    let i = base + 1;
    let len = base + 2;

    emit_is_go_error_slot(&mut chunks[current], cur, line);
    chunks[current].emit_if(line);
    emit_get_field_slot(&mut chunks[current], cur, "errs", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, errs, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, errs, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    emit_array_len_slot(&mut chunks[current], errs, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    let outer = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, work, line);
    emit_array_get_slot(chunks, current, errs, i, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_end(line);

    emit_get_field_slot(&mut chunks[current], cur, "wrap", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, errs, line);
    emit_push_local_if_non_null(chunks, current, work, errs, line);
    chunks[current].emit_end(line);
}

fn emit_push_local_if_non_null(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    value_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn emit_is_go_error_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    emit_get_field_slot(chunk, slot, GO_ERROR_FLAG, line);
    ops::emit_dyn_to_bool(chunk, line);
}

fn emit_array_len_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let len = chunk.add_import("ecma:array", "length");
    chunk.emit_call(len, 1, line);
}

fn emit_array_get_slot(chunks: &mut [Chunk], current: usize, array: u16, index: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, array, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_get(chunks, current, line);
}

fn emit_bool_from_i32_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

fn emit_get_field_slot(chunk: &mut Chunk, obj: u16, key: &str, line: u32) {
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::internal(key),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Local(obj),
        &slot,
        class_slots::Dest::Stack,
        line,
    );
}

fn set_field_local(chunk: &mut Chunk, obj: u16, key: &str, value: u16, line: u32) {
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::internal(key),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Local(obj),
        &slot,
        class_slots::ValueSource::Local(value),
        line,
    );
}

fn set_field_str(chunk: &mut Chunk, obj: u16, key: &str, value: &str, line: u32) {
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::internal(key),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Local(obj),
        &slot,
        class_slots::ValueSource::ConstStr(value.to_string()),
        line,
    );
}

fn set_field_bool(chunk: &mut Chunk, obj: u16, key: &str, value: bool, line: u32) {
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::internal(key),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Local(obj),
        &slot,
        class_slots::ValueSource::ConstBool(value),
        line,
    );
}
