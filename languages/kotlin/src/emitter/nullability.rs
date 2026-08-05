//! Kotlin nullability operators.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn emit_exception_throw(chunk: &mut Chunk, exc_name: &str, message_slot: Option<u16>, line: u32) {
    chunk.emit_struct_new(0, 0, line);
    chunk.emit_dup(line);
    if let Some(slot) = message_slot {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        chunk.emit_op_u8(Op::CALL_REF, 0, line);
    } else {
        chunk.emit_string_const(exc_name, line);
    }
    vybe_compiler::primitives::errors::emit_exception_new_finalize(chunk, exc_name, line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

/// Kotlin/JVM exception constructor.
///
/// Stack: `[message?]` -> `[exception_object]`.
pub fn emit_exception(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    exc_name: &str,
    line: u32,
) {
    let msg = chunks[current].alloc_scratch(1);
    if argc > 0 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, msg, line);
    } else {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, msg, line);
    }

    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, msg, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        exc_name,
        line,
    );
}

/// Kotlin `x!!`.
///
/// Stack: `[value]` -> `[value]`, or throws `NullPointerException` when value is null.
pub fn emit_not_null_assert(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }

    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_exception_throw(&mut chunks[current], "NullPointerException", None, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_end(line);
}

/// Kotlin `require(condition) { message }` / `check(condition) { message }`.
///
/// Stack: `[condition]` or `[condition, message_lambda]` -> `[null]`, or throws.
pub fn emit_precondition(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    exc_name: &str,
    line: u32,
) {
    let message = if argc >= 2 {
        let slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
        Some(slot)
    } else {
        None
    };
    let condition = chunks[current].alloc_scratch(1);
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, condition, line);
    } else {
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, condition, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, condition, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    emit_exception_throw(&mut chunks[current], exc_name, message, line);
    chunks[current].emit_end(line);
}

/// Kotlin `requireNotNull(value)` / `checkNotNull(value)`.
///
/// Stack: `[value]` -> `[value]`, or throws the given precondition exception.
pub fn emit_precondition_not_null(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    exc_name: &str,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_exception_throw(&mut chunks[current], exc_name, None, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_end(line);
}

/// Kotlin value reflection: `value::class`.
///
/// Stack: `[value]` -> `[{ simpleName, qualifiedName, java }]`.
pub fn emit_class_of(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    }

    let type_name = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    let type_key = chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
        vybe_compiler::primitives::reflection::FIELD_TYPE_NAME,
    )));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_name, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, type_name, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("Any", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_name, line);
    chunks[current].emit_end(line);

    let klass = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, klass, line);
    set_field_from_slot(&mut chunks[current], klass, "simpleName", type_name, line);
    set_field_from_slot(
        &mut chunks[current],
        klass,
        "qualifiedName",
        type_name,
        line,
    );

    let java = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, java, line);
    set_field_from_slot(&mut chunks[current], java, "name", type_name, line);
    set_field_from_slot(&mut chunks[current], java, "canonicalName", type_name, line);
    set_field_from_slot(&mut chunks[current], java, "simpleName", type_name, line);
    set_field_from_slot(&mut chunks[current], klass, "java", java, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, klass, line);
}

pub fn emit_error(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_exception(chunks, current, argc, "IllegalStateException", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

fn set_field_from_slot(
    chunk: &mut Chunk,
    object_slot: u16,
    field: &str,
    value_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let key = chunk.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(field)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

/// `x ?: throw e` — the throw is an EXPRESSION here; the helper throws the
/// exception object on the stack (the trailing null only balances types —
/// it is unreachable).
pub fn emit_throw_expr(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
