//! Java `Optional` adapters.
//!
//! The shared host has no `ecma:optional` module, so Java lowers Optional to a
//! tiny pair array: `[present: bool, value]`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::strings;

pub fn emit_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    collections::emit_array_new(chunks, current, 2, line);
}

pub fn emit_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_array_new(chunks, current, 2, line);
}

pub fn emit_of_long(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_string_const("long", line);
    collections::emit_array_new(chunks, current, 3, line);
}

pub fn emit_of_nullable(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_is_present(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
}

pub fn emit_or_else(chunks: &mut [Chunk], current: usize, call_supplier: bool, line: u32) {
    let fallback_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fallback_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fallback_slot, line);
    if call_supplier {
        chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_if_present(chunks: &mut [Chunk], current: usize, line: u32) {
    let consumer_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, consumer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, consumer_slot, line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_filter(chunks: &mut [Chunk], current: usize, line: u32) {
    let predicate_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, predicate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, predicate_slot, line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let mapper_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapper_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mapper_slot, line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    emit_of_nullable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_flat_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let mapper_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapper_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mapper_slot, line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_if_present_or_else(chunks: &mut [Chunk], current: usize, line: u32) {
    let empty_action_slot = chunks[current].alloc_scratch(1);
    let consumer_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, empty_action_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, consumer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, consumer_slot, line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, empty_action_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_or(chunks: &mut [Chunk], current: usize, call_supplier: bool, line: u32) {
    let fallback_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fallback_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fallback_slot, line);
    if call_supplier {
        chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_stream(chunks: &mut [Chunk], current: usize, line: u32) {
    let optional_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let other_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::object::emit_equals(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_tag_from_slot(chunks, current, optional_slot, line);
    emit_tag_from_slot(chunks, current, other_slot, line);
    vybe_compiler::primitives::object::emit_equals(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    emit_value_from_slot(chunks, current, other_slot, line);
    vybe_compiler::primitives::object::emit_equals(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let optional_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("Optional[", line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    let to_string = chunks[current].add_import("ecma:string", "String");
    chunks[current].emit_call(to_string, 1, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_string_const("]", line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("Optional.empty", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_or_else_throw(chunks: &mut [Chunk], current: usize, has_supplier: bool, line: u32) {
    let supplier_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    if has_supplier {
        chunks[current].emit_op_u16(Op::LOCAL_SET, supplier_slot, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, supplier_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    if has_supplier {
        chunks[current].emit_op_u16(Op::LOCAL_GET, supplier_slot, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    } else {
        chunks[current].emit_struct_new(0, 0, line);
        chunks[current].emit_dup(line);
        chunks[current].emit_string_const("", line);
        vybe_compiler::primitives::errors::emit_exception_new_finalize(
            &mut chunks[current],
            "java.util.NoSuchElementException",
            line,
        );
    }
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_value_from_slot(chunks: &mut [Chunk], current: usize, optional_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
}

fn emit_tag_from_slot(chunks: &mut [Chunk], current: usize, optional_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    chunks[current].emit_i32_const(2, line);
    collections::emit_get(chunks, current, line);
}
