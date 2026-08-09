//! Shared .NET-shaped helpers routed through runtime helper chunks.

use std::sync::Arc;

use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::loops;
use vybe_runtime::Chunk;
use vybe_runtime::Op;
use vybe_runtime::Value;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    if name == "dotnet.tostring" {
        let to_str = chunks[current].add_import("ecma:string", "String");
        chunks[current].emit_call(to_str, argc, line);
        return true;
    }

    if name == "dotnet.tostring_runtime" {
        emit_tostring_runtime(&mut chunks[current], argc, line);
        return true;
    }

    if name == "dotnet.string_join_sep_first" {
        emit_string_join_sep_first(chunks, current, line);
        return true;
    }

    if name == "dotnet.sort_with_comparator" {
        collections::emit_sort_with_comparator(chunks, current, line);
        return true;
    }

    if name == "dotnet.string_is_null_or_empty" {
        collections::emit_runtime_helper_call(
            chunks,
            current,
            "__vybe_string_is_null_or_empty",
            argc,
            line,
        );
        vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
        return true;
    }

    if name == "dotnet.string_is_null_or_whitespace" {
        collections::emit_runtime_helper_call(
            chunks,
            current,
            "__vybe_string_is_null_or_whitespace",
            argc,
            line,
        );
        vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
        return true;
    }

    if name == "dotnet.cchar" {
        emit_cchar(&mut chunks[current], line);
        return true;
    }

    let global = match name {
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

// `char` → code point moved OUT of this adapter to
// `vybe_compiler::primitives::strings::emit_char_code`, reachable as
// `common:strings.char_code`. It is a number conversion, not a .NET surface, and
// the languages bind it through the `char`/`int` coercion slot.

/// .NET `CChar` / `Convert.ToChar`.
///
/// Numeric inputs are Unicode code points; string inputs return the first
/// UTF-16 code unit as a one-character string. Keeping this in the dotnet
/// adapter avoids VB-only coercion rules and lets every dotnet language share
/// the same conversion semantics.
fn emit_cchar(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(3);
    let ty = value + 1;
    let result = value + 2;

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ty, line);

    chunk.emit_op_u16(Op::LOCAL_GET, ty, line);
    chunk.emit_string_const("string", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_i32_const(1, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(chunk, "ecma:string", "fromCharCode", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
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
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, is_numeric, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, is_numeric, line);
    chunk.emit_if_value(line);
    // Numeric: render via the shared `__vybe_dotnet_numeric_format(v, fmt, 0)`.
    vybe_compiler::primitives::globals::emit_read(chunk, "__vybe_dotnet_numeric_format", line);
    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fmt, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 3, 1, line);
    chunk.emit_else(line);
    // Non-numeric: the format is a no-op, take the ordinary dispatch.
    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_tostring_dispatch(chunk, line);
    chunk.emit_end(line);
}

/// Runtime `ToString()` type dispatch. Stack: `[obj]` → `[string]`. Mirrors the
/// retired `zero_arg_tostring` fallback: primitives and objects with no
/// `ToString` role go through `String()`; an object carrying the shared role
/// method calls it; a Guid struct (`__type == "Guid"`) renders its `__value`.
fn emit_tostring_dispatch(chunk: &mut Chunk, line: u32) {
    let obj = chunk.alloc_scratch(1);
    let ty = chunk.alloc_scratch(1);
    let result = chunk.alloc_scratch(1);
    let is_primitive = chunk.alloc_scratch(1);
    let func = chunk.alloc_scratch(1);

    let tostring_key = chunk.add_constant(Value::String(Arc::from(vybe_ast::protocol_slot_key(
        vybe_ast::ProtocolSlot::ToString,
    ))));
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    let value_key = chunk.add_constant(Value::String(Arc::from("__value")));
    let buf_key = chunk.add_constant(Value::String(Arc::from("__buf")));

    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ty, line);

    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, is_primitive, line);

    for type_name in ["number", "i32", "i64", "string", "boolean"] {
        chunk.emit_op_u16(Op::LOCAL_GET, is_primitive, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, ty, line);
        chunk.emit_string_const(type_name, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, is_primitive, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, is_primitive, line);
    chunk.emit_if(line);
    // Primitive: .NET stringification. Booleans are capitalized (`True` /
    // `False`), other primitives follow ECMA String().
    chunk.emit_op_u16(Op::LOCAL_GET, ty, line);
    chunk.emit_string_const("boolean", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_string_const("True", line);
    chunk.emit_else(line);
    chunk.emit_string_const("False", line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    vybe_compiler::primitives::strings::emit_to_string(chunk, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_else(line);

    // Object: look up its shared `ToString` role member.
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, tostring_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, func, line);

    chunk.emit_op_u16(Op::LOCAL_GET, func, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_if(line);
    // No `tostring` member: known .NET-shaped structs/classes render their
    // payload, otherwise fall back to ECMA String().
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    chunk.emit_string_const("Guid", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    chunk.emit_string_const("StringWriter", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    vybe_compiler::primitives::strings::emit_to_string(chunk, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_else(line);
    // Has a `ToString` role member: call it with the receiver.
    chunk.emit_op_u16(Op::LOCAL_GET, func, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
}

/// .NET `String.Join(separator, values)`.
///
/// ECMA `Array.join` stringifies with JavaScript semantics (`true`,
/// `[object]`). .NET goes through each element's `ToString` role first, then
/// joins the materialized strings. Stack: `[separator, values] -> [string]`.
fn emit_string_join_sep_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let sep_slot = chunks[current].alloc_scratch(5);
    let values_slot = sep_slot + 1;
    let mapped_slot = sep_slot + 2;
    let idx_slot = sep_slot + 3;
    let elem_slot = sep_slot + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, values_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sep_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_slot, line);

    let state = loops::emit_for_in_start(chunks, current, values_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    emit_tostring_dispatch(&mut chunks[current], line);
    chunks[current].emit_end(line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "join", 2, line);
}
