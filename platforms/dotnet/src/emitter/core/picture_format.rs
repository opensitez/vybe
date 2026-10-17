//! VB/.NET picture formatting, emitted once per module by the picture adapter.
//! Reusable string and numeric operations remain shared compiler primitives.
use vybe_runtime::{Chunk, opcode::Op};
use vybe_compiler::primitives::strings::{
    emit_length as emit_str_length, emit_substring as emit_str_substring,
    emit_str_char_code_at, emit_str_equals, emit_str_concat,
};

pub fn build_picture_format() -> Chunk {
    let mut c = Chunk::new("__dotnet_vb_picture_format");
    c.arity = 2;
    c.local_count = 12; // value(0), picture(1), prefix(2), dot_pos(3), decimals(4), fmt_lower(5), val_str(6), work(7), idx_a(8), idx_b(9), idx_c(10), percent(11)
    let value = 0u16;
    let picture = 1u16;
    let prefix = 2u16;
    let dot_pos = 3u16;
    let decimals = 4u16;
    let fmt_lower = 5u16;
    let val_str = 6u16;
    let work = 7u16;
    let idx_a = 8u16;
    let idx_b = 9u16;
    let idx_c = 10u16;
    let percent = 11u16;

    let to_str = c.add_import("ecma:string", "String");
    let to_lower = c.add_import("ecma:string", "toLowerCase");
    let pad_start = c.add_import("ecma:string", "padStart");
    let to_fixed = c.add_import("ecma:number", "toFixed");
    let parse_int = c.add_import("ecma:number", "parseInt");

    // prefix = ""
    let empty = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("")));
    let short_date = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
        "short date",
    )));
    let short_time = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
        "short time",
    )));
    let percent_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("%")));
    let zero_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("0")));
    let space_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(" ")));
    let slash_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("/")));
    let colon_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(":")));
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, empty, 0);
    c.emit_op_u16(Op::LOCAL_SET, prefix, 0);
    c.emit_bool_const(false, 0);
    c.emit_op_u16(Op::LOCAL_SET, percent, 0);

    // If picture is null/empty, return String(value).
    let no_picture_block = c.emit_block(0);
    let null_or_empty = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let picture_null = c.add_constant(vybe_runtime::Value::I32(0));
    let _ = picture_null; // not used; kept for future readability
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    // null path → return String(value)
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(null_or_empty);

    // Check empty string ("" length = 0)
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(no_picture_block);

    // fmt_lower = picture.Trim().ToLowerCase(); val_str = String(value)
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    {
        let idx = c.add_import("ecma:string", "trim");
        c.emit_call(idx, 1, 0);
    }
    c.emit_call(to_lower, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, fmt_lower, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, val_str, 0);

    // Short Date → first segment before space, else whole value string.
    let not_short_date = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, fmt_lower, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, short_date, 0);
    emit_str_equals(&mut c, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, space_str, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, idx_a, 0);
    let no_date_space = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    c.emit_f64_const(0.0, 0);
    c.emit_op(Op::F64_LT, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(no_date_space);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(not_short_date);

    // Short Time → trim optional date prefix, drop seconds, keep AM/PM.
    let not_short_time = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, fmt_lower, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, short_time, 0);
    emit_str_equals(&mut c, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op_u16(Op::LOCAL_TEE, work, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, space_str, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, idx_a, 0);
    let no_prefix = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    c.emit_f64_const(0.0, 0);
    c.emit_op(Op::F64_LT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_TEE, idx_b, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, slash_str, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, idx_c, 0);
    let keep_work = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, idx_c, 0);
    c.emit_f64_const(0.0, 0);
    c.emit_op(Op::F64_LT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    emit_str_length(&mut c, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_end(0);
    c.patch_block(keep_work);
    c.emit_end(0);
    c.patch_block(no_prefix);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, space_str, 0);
    {
        let idx = c.add_import("ecma:string", "lastIndexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, idx_a, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, idx_b, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    emit_str_length(&mut c, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, idx_c, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_b, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, colon_str, 0);
    {
        let idx = c.add_import("ecma:string", "lastIndexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, idx_a, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_b, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, colon_str, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    let already_short_time = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    c.emit_op(Op::F64_EQ, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_b, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, space_str, 0);
    emit_str_concat(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_c, 0);
    emit_str_concat(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(already_short_time);
    c.emit_op_u16(Op::LOCAL_GET, idx_b, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    emit_str_substring(&mut c, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, space_str, 0);
    emit_str_concat(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_c, 0);
    emit_str_concat(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(not_short_time);

    // If picture starts with '$', strip it and stash as prefix.
    let dollar_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    emit_str_char_code_at(&mut c, 0);
    let dollar_code = c.add_constant(vybe_runtime::Value::I32(b'$' as i32));
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, dollar_code, 0);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    // prefix = "$"
    let dollar_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("$")));
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, dollar_str, 0);
    c.emit_op_u16(Op::LOCAL_SET, prefix, 0);
    // picture = picture.substring(1)
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, picture, 0);
    c.emit_end(0);
    c.patch_block(dollar_block);

    // If picture ends with '%', strip it and mark percentage mode.
    let percent_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op(Op::I32_GT_S, 0);
    c.emit_op(vybe_runtime::opcode::Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    emit_str_char_code_at(&mut c, 0);
    let percent_code = c.add_constant(vybe_runtime::Value::I32(b'%' as i32));
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, percent_code, 0);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_bool_const(true, 0);
    c.emit_op_u16(Op::LOCAL_SET, percent, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, picture, 0);
    c.emit_end(0);
    c.patch_block(percent_block);

    // dot_pos = picture.indexOf(".")
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    let dot_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(".")));
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, dot_str, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, dot_pos, 0);

    // If no dot: return prefix + String(parseInt(value))
    let no_decimals_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, dot_pos, 0);
    c.emit_f64_const(0.0, 0);
    c.emit_op(Op::F64_LT, 0);
    c.emit_op(vybe_runtime::opcode::Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    // No dot — integer rendering, optionally zero-padded / percentage.
    let no_pct_int = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, percent, 0);
    let cast = c.add_import("wasm:js-boolean", "cast");
    c.emit_call(cast, 1, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    let hundred = c.add_constant(vybe_runtime::Value::F64(100.0));
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, hundred, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_end(0);
    c.emit_end(0);
    c.patch_block(no_pct_int);

    c.emit_op_u16(Op::LOCAL_GET, prefix, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_call(parse_int, 1, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    let zero_pad_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_GT_S, 0);
    c.emit_op(vybe_runtime::opcode::Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    emit_str_char_code_at(&mut c, 0);
    let zero_code = c.add_constant(vybe_runtime::Value::I32(b'0' as i32));
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, zero_code, 0);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, zero_str, 0);
    c.emit_call(pad_start, 3, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_end(0);
    c.patch_block(zero_pad_block);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    emit_str_concat(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, val_str, 0);
    let no_pct_suffix_int = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, percent, 0);
    let cast = c.add_import("wasm:js-boolean", "cast");
    c.emit_call(cast, 1, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, percent_str, 0);
    emit_str_concat(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, val_str, 0);
    c.emit_end(0);
    c.patch_block(no_pct_suffix_int);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(no_decimals_block);

    // decimals = picture.length - dot_pos - 1
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, dot_pos, 0);
    c.emit_op(Op::I32_SUB, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, decimals, 0);

    // return prefix + Number(value).toFixed(decimals), with optional percentage suffix
    c.emit_op_u16(Op::LOCAL_GET, prefix, 0);
    let no_pct_dec = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, percent, 0);
    let cast = c.add_import("wasm:js-boolean", "cast");
    c.emit_call(cast, 1, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    let hundred_dec = c.add_constant(vybe_runtime::Value::F64(100.0));
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, hundred_dec, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_end(0);
    c.emit_end(0);
    c.patch_block(no_pct_dec);

    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_op_u16(Op::LOCAL_GET, decimals, 0);
    c.emit_call(to_fixed, 2, 0);
    emit_str_concat(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, val_str, 0);
    let no_pct_suffix_dec = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, percent, 0);
    let cast = c.add_import("wasm:js-boolean", "cast");
    c.emit_call(cast, 1, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, percent_str, 0);
    emit_str_concat(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, val_str, 0);
    c.emit_end(0);
    c.patch_block(no_pct_suffix_dec);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

