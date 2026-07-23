//! Shared .NET-shaped helpers routed through runtime helper chunks.

use std::sync::Arc;

use vybe_bytecode::Chunk;
use vybe_bytecode::Op;
use vybe_bytecode::Value;
use vybe_emitter::collections;
use vybe_emitter::instructions::host;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    if name == "dotnet.tostring" {
        let to_str = chunks[current].add_import("ecma:string", "String");
        chunks[current].emit_op_u16(Op::CALL_IMPORT, to_str, line);
        chunks[current].emit(argc, line);
        return true;
    }

    if name == "dotnet.tostring_runtime" {
        emit_tostring_runtime(&mut chunks[current], argc, line);
        return true;
    }

    if name == "dotnet.sort_with_comparator" {
        collections::emit_sort_with_comparator(chunks, current, line);
        return true;
    }

    let global = match name {
        "dotnet.cchar" => "__vybe_cchar",
        "dotnet.string_is_null_or_empty" => "__vybe_string_is_null_or_empty",
        "dotnet.string_is_null_or_whitespace" => "__vybe_string_is_null_or_whitespace",
        "dotnet.newline" => "__vybe_newline",
        "dotnet.str_insert" => "__vybe_str_insert",
        "dotnet.str_remove_start" => "__vybe_str_remove_start",
        "dotnet.str_remove_range" => "__vybe_str_remove_range",
        "dotnet.sort_in_place" => "__vybe_sort_in_place",
        "dotnet.val" => "__vybe_val",
        "dotnet.iif" => "__vybe_iif",
        "dotnet.rgb" => "__vybe_rgb",
        "dotnet.qbcolor" => "__vybe_qbcolor",
        "dotnet.isnumeric" => "__vybe_isnumeric",
        "dotnet.isempty" => "__vybe_isempty",
        "dotnet.isdate" => "__vybe_isdate",
        "dotnet.vartype" => "__vybe_vartype",
        "dotnet.regex_match_all_pat_first" => "__ecma_regexp_match_all_pat_first",
        "dotnet.regex_replace_pat_first" => "__ecma_regexp_replace_pat_first",
        "dotnet.regex_split_pat_first" => "__ecma_regexp_split_pat_first",
        "dotnet.array_remove_value" => "__vybe_array_remove_value",
        "dotnet.array_remove_at" => "__vybe_array_remove_at",
        "dotnet.array_insert" => "__vybe_array_insert",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}

/// Runtime `ToString([format])` dispatch for receivers whose type is unknown at
/// compile time (typed receivers resolve through the surface path instead).
/// `argc` counts the receiver plus any format argument.
///
/// - `ToString()` (argc 1): runtime type dispatch (`emit_tostring_dispatch`).
/// - `ToString(fmt)` (argc 2): a numeric receiver renders through the shared
///   `__vybe_dotnet_numeric_format` helper (D/X/F/N/P/… specifiers); any other
///   receiver ignores the format and takes the plain dispatch (the numeric
///   renderer would otherwise yield "nan" for it).
fn emit_tostring_runtime(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc < 2 {
        emit_tostring_dispatch(chunk, line);
        return;
    }

    // Stack: [receiver, format].
    let fmt = chunk.alloc_scratch(1);
    let recv = chunk.alloc_scratch(1);
    let ty = chunk.alloc_scratch(1);
    let is_numeric = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, fmt, line);
    chunk.emit_op_u16(Op::LOCAL_SET, recv, line);

    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ty, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, is_numeric, line);
    for type_name in ["number", "i32", "i64", "f64", "f32"] {
        chunk.emit_op_u16(Op::LOCAL_GET, is_numeric, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, ty, line);
        chunk.emit_string_const(type_name, line);
        vybe_emitter::ops::emit_dyn_eq(chunk, line);
        vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, is_numeric, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, is_numeric, line);
    chunk.emit_if_value(line);
    // Numeric: render via the shared `__vybe_dotnet_numeric_format(v, fmt, 0)`.
    let helper = chunk.add_constant(Value::String(Arc::from("__vybe_dotnet_numeric_format")));
    chunk.emit_op_u16(Op::GLOBAL_GET, helper, line);
    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fmt, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(3, line);
    chunk.emit_else(line);
    // Non-numeric: the format is a no-op, take the ordinary dispatch.
    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_tostring_dispatch(chunk, line);
    chunk.emit_end(line);
}

/// Runtime `ToString()` type dispatch. Stack: `[obj]` → `[string]`. Mirrors the
/// retired `zero_arg_tostring` fallback: primitives and objects with no
/// `tostring` member go through `String()`; an object carrying a `tostring`
/// method calls it; a Guid struct (`__type == "Guid"`) renders its `__value`.
fn emit_tostring_dispatch(chunk: &mut Chunk, line: u32) {
    let obj = chunk.alloc_scratch(1);
    let ty = chunk.alloc_scratch(1);
    let result = chunk.alloc_scratch(1);
    let is_primitive = chunk.alloc_scratch(1);
    let func = chunk.alloc_scratch(1);

    let tostring_key = chunk.add_constant(Value::String(Arc::from("tostring")));
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    let value_key = chunk.add_constant(Value::String(Arc::from("__value")));

    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ty, line);

    chunk.emit_op(Op::NULL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, is_primitive, line);

    for type_name in ["number", "i32", "i64", "string", "boolean"] {
        chunk.emit_op_u16(Op::LOCAL_GET, is_primitive, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, ty, line);
        chunk.emit_string_const(type_name, line);
        vybe_emitter::ops::emit_dyn_eq(chunk, line);
        vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, is_primitive, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, is_primitive, line);
    chunk.emit_if(line);
    // Primitive: plain String() coercion.
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    vybe_emitter::strings::emit_to_string(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_else(line);

    // Object: look up its `tostring` member.
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::STRUCT_GET, tostring_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, func, line);

    chunk.emit_op_u16(Op::LOCAL_GET, func, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_if(line);
    // No `tostring` member: a Guid renders its stored value, else String().
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::STRUCT_GET, type_key, line);
    chunk.emit_string_const("Guid", line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::STRUCT_GET, value_key, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    vybe_emitter::strings::emit_to_string(chunk, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_else(line);
    // Has a `tostring` member: call it with the receiver.
    chunk.emit_op_u16(Op::LOCAL_GET, func, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
}
