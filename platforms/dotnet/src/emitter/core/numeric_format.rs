//! `.NET` numeric format-string rendering — `value.ToString(format[, width])`.
//!
//! Builds the `__stdlib_dotnet_numeric_format` runtime-helper chunk that
//! renders a number against a .NET standard/custom numeric format specifier
//! (`D`, `X`, `F`, `P`, optional precision digits) plus an optional field
//! width. Declared ONCE here in the platform; every dotnet-shaped language
//! (C#, VB, …) reaches it through the common resolver, so the formatting is
//! never reimplemented per grammar.
//!
//! Relocated out of the shared compiler (`vybe_compiler::primitives::
//! runtime_helpers`) — this is dotnet-specific codegen and belongs in the
//! dotnet platform. The shared bundler calls [`build_dotnet_numeric_format`]
//! only to include the chunk when a program uses it.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::ops::{
    emit_dyn_add_into, emit_dyn_eq_into, emit_dyn_ge_into, emit_dyn_gt_into, emit_dyn_lt_into,
    emit_dyn_not_into,
};

fn emit_str_length(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(idx, 1, line);
}

fn emit_str_substring(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "substring");
    chunk.emit_call(idx, 3, line);
}

fn emit_str_char_code_at(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "charCodeAt");
    chunk.emit_call(idx, 2, line);
}

fn emit_str_concat(imports: &mut Chunk, chunk: &mut Chunk, line: u32) {
    emit_dyn_add_into(imports, chunk, line);
}

fn emit_const_index(chunk: &mut Chunk, idx: u16, line: u32) {
    match chunk.constants[idx as usize].clone() {
        Value::Null | Value::TypedNull(_) => chunk.emit_op(Op::NULL, line),
        Value::Undefined => core_wasm::undefined(chunk, line),
        Value::Bool(value) => chunk.emit_bool_const(value, line),
        Value::I32(value) => chunk.emit_i32_const(value, line),
        Value::I64(value) => chunk.emit_i64_const(value, line),
        Value::BigInt(value) => chunk.emit_i64_const(value.to_i64_wrapping(), line),
        Value::F64(value) => chunk.emit_f64_const(value, line),
        Value::F32(value) => chunk.emit_f32_const(value, line),
        Value::String(value) | Value::Symbol(value) => chunk.emit_string_const(&value, line),
        Value::Object(_) | Value::WeakRef(_) | Value::V128(_) => {
            panic!("runtime helper cannot inline non-primitive constant")
        }
    }
}

pub fn build_dotnet_numeric_format(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_dotnet_numeric_format");
    c.arity = 3;
    c.local_count = 8;
    let value = 0u16;
    let format = 1u16;
    let width = 2u16;
    let fmt = 3u16;
    let precision = 4u16;
    let first_code = 5u16;
    let rendered = 6u16;
    let abs_width = 7u16;

    let to_str = c.add_import("ecma:string", "String");
    let parse_int = c.add_import("ecma:number", "parseInt");
    let number = c.add_import("ecma:number", "Number");
    let number_to_string = c.add_import("ecma:number", "toString");
    let to_fixed = c.add_import("ecma:number", "toFixed");
    let to_upper = c.add_import("ecma:string", "toUpperCase");
    let pad_start = c.add_import("ecma:string", "padStart");
    let pad_end = c.add_import("ecma:string", "padEnd");

    let zero_num = c.add_constant(Value::F64(0.0));
    let sixteen = c.add_constant(Value::F64(16.0));
    let hundred = c.add_constant(Value::F64(100.0));
    let zero_str = c.add_constant(Value::String(Arc::from("0")));
    let space_str = c.add_constant(Value::String(Arc::from(" ")));
    let minus_str = c.add_constant(Value::String(Arc::from("-")));
    let percent_suffix = c.add_constant(Value::String(Arc::from(" %")));
    let d_code = c.add_constant(Value::I32(b'D' as i32));
    let x_code = c.add_constant(Value::I32(b'X' as i32));
    let f_code = c.add_constant(Value::I32(b'F' as i32));
    let p_code = c.add_constant(Value::I32(b'P' as i32));
    let minus_code = c.add_constant(Value::I32(b'-' as i32));

    let has_format = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, format, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(has_format);

    c.emit_op_u16(Op::LOCAL_GET, format, 0);
    c.emit_call(to_str, 1, 0);
    {
        let idx = c.add_import("ecma:string", "trim");
        c.emit_call(idx, 1, 0);
    }
    c.emit_call(to_upper, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, fmt, 0);

    let non_empty_format = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, fmt, 0);
    emit_str_length(&mut c, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    emit_dyn_eq_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(non_empty_format);

    emit_const_index(&mut c, zero_num, 0);
    c.emit_op_u16(Op::LOCAL_SET, precision, 0);

    let no_precision_suffix = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, fmt, 0);
    emit_str_length(&mut c, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    emit_dyn_gt_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, fmt, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_GET, fmt, 0);
    emit_str_length(&mut c, 0);
    emit_str_substring(&mut c, 0);
    c.emit_call(parse_int, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, precision, 0);
    let precision_is_number = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    emit_const_index(&mut c, zero_num, 0);
    c.emit_op_u16(Op::LOCAL_SET, precision, 0);
    c.emit_end(0);
    c.patch_block(precision_is_number);
    c.emit_end(0);
    c.patch_block(no_precision_suffix);

    c.emit_op_u16(Op::LOCAL_GET, fmt, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    emit_str_char_code_at(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, first_code, 0);

    let dispatch = c.emit_block(0);

    let not_decimal = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, first_code, 0);
    emit_const_index(&mut c, d_code, 0);
    emit_dyn_eq_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(parse_int, 1, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);

    let non_negative = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    emit_str_length(&mut c, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    emit_dyn_gt_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    emit_str_char_code_at(&mut c, 0);
    emit_const_index(&mut c, minus_code, 0);
    emit_dyn_eq_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    emit_const_index(&mut c, minus_str, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    emit_str_length(&mut c, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    emit_const_index(&mut c, zero_str, 0);
    c.emit_call(pad_start, 3, 0);
    emit_str_concat(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_br(2, 0);
    c.emit_end(0);
    c.patch_block(non_negative);

    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    emit_const_index(&mut c, zero_str, 0);
    c.emit_call(pad_start, 3, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(not_decimal);

    let not_hex = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, first_code, 0);
    emit_const_index(&mut c, x_code, 0);
    emit_dyn_eq_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(number, 1, 0);
    emit_const_index(&mut c, sixteen, 0);
    c.emit_call(number_to_string, 2, 0);
    c.emit_call(to_upper, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    emit_const_index(&mut c, zero_str, 0);
    c.emit_call(pad_start, 3, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(not_hex);

    let not_fixed = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, first_code, 0);
    emit_const_index(&mut c, f_code, 0);
    emit_dyn_eq_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(number, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    c.emit_call(to_fixed, 2, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(not_fixed);

    let not_percent = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, first_code, 0);
    emit_const_index(&mut c, p_code, 0);
    emit_dyn_eq_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(number, 1, 0);
    emit_const_index(&mut c, hundred, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    c.emit_call(to_fixed, 2, 0);
    emit_const_index(&mut c, percent_suffix, 0);
    emit_str_concat(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(not_percent);

    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);

    c.emit_end(0);
    c.patch_block(dispatch);

    let width_is_number = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(width_is_number);

    let width_is_zero = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    emit_const_index(&mut c, zero_num, 0);
    emit_dyn_eq_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(width_is_zero);

    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    c.emit_op_u16(Op::LOCAL_SET, abs_width, 0);
    let width_non_negative = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    emit_const_index(&mut c, zero_num, 0);
    emit_dyn_lt_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    emit_const_index(&mut c, zero_num, 0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    c.emit_op(Op::F64_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, abs_width, 0);
    c.emit_end(0);
    c.patch_block(width_non_negative);

    let already_wide_enough = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    emit_str_length(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, abs_width, 0);
    emit_dyn_ge_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(already_wide_enough);

    let right_aligned = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    emit_const_index(&mut c, zero_num, 0);
    emit_dyn_lt_into(imports, &mut c, 0);
    emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op_u16(Op::LOCAL_GET, abs_width, 0);
    emit_const_index(&mut c, space_str, 0);
    c.emit_call(pad_end, 3, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(right_aligned);

    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op_u16(Op::LOCAL_GET, abs_width, 0);
    emit_const_index(&mut c, space_str, 0);
    c.emit_call(pad_start, 3, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
