//! Pascal runtime-surface helpers routed via `common:pascal.*`.

use vybe_bytecode::Chunk;
use vybe_bytecode::Op;
use vybe_bytecode::Value;
use vybe_compiler::compiler::collections;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    if name == "pascal.tostring" {
        let to_str = chunks[0].add_import("ecma:string", "String");
        chunks[current].emit_op_u16(Op::CALL_IMPORT, to_str, line);
        chunks[current].emit(argc, line);
        return true;
    }

    if name == "pascal.ord" {
        emit_ord(chunks, current, line);
        return true;
    }

    if name == "pascal.set_length" {
        let idx = chunks[0].add_import("ecma:set", "size");
        chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
        chunks[current].emit(argc, line);
        return true;
    }

    if name == "pascal.int_xor" {
        chunks[current].emit_op(Op::I32_XOR, line);
        return true;
    }

    if name == "pascal.file_eof" {
        emit_pascal_file_eof(chunks, current, line);
        return true;
    }

    let global = match name {
        "pascal.str_remove_range" => "__vybe_str_remove_range",
        "pascal.str_insert" => "__vybe_str_insert",
        "pascal.sort_in_place" => "__vybe_sort_in_place",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}

fn emit_pascal_file_eof(chunks: &mut [Chunk], current: usize, line: u32) {
    let handle_slot = chunks[current].alloc_scratch(1);
    let next_slot = chunks[current].alloc_scratch(1);
    let rows_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let eof_map_slot = ensure_global_map(chunks, current, "__vb_file_eof_by_handle", line);
    let next_map_slot =
        ensure_global_map(chunks, current, "__vb_record_next_index_by_handle", line);
    let rows_map_slot = ensure_global_map(chunks, current, "__vb_record_rows_by_handle", line);

    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, handle_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, next_map_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_SET, next_slot, line);

        chunk.emit_op_u16(Op::LOCAL_GET, rows_map_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_SET, rows_slot, line);

        chunk.emit_op_u16(Op::LOCAL_GET, next_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, eof_map_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        vybe_compiler::compiler::instructions::core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }
    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, eof_map_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        vybe_compiler::compiler::instructions::core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
    }
    collections::emit_len(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, next_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
        vybe_compiler::compiler::ops::emit_dyn_ge(chunk, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn ensure_global_map(chunks: &mut [Chunk], current: usize, name: &str, line: u32) -> u16 {
    let slot = chunks[current].alloc_scratch(1);
    let key = chunks[current].add_constant(Value::String(std::sync::Arc::from(name)));
    chunks[current].emit_op_u16(Op::GLOBAL_GET, key, line);
    vybe_compiler::compiler::instructions::core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    collections::emit_map_new(chunks, current, line);
    vybe_compiler::compiler::instructions::core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::GLOBAL_SET, key, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn emit_ord(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.local_count;
    let result_slot = value_slot + 1;
    chunk.alloc_scratch(2);

    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    let str_test = chunk.add_import("wasm:js-string", "test");
    let char_code_at = chunk.add_import("ecma:string", "charCodeAt");
    let bool_test = chunk.add_import("wasm:js-boolean", "test");
    let bool_cast = chunk.add_import("wasm:js-boolean", "cast");
    let num_to_f64 = chunk.add_import("wasm:js-number", "toF64");

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(str_test, 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(char_code_at, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(bool_test, 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(bool_cast, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(num_to_f64, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}
