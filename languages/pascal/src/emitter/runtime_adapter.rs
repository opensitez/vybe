//! Pascal runtime-surface helpers routed via `common:pascal.*`.

use vybe_bytecode::Chunk;
use vybe_bytecode::Op;
use vybe_emitter::collections;

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

    let global = match name {
        "pascal.str_remove_range" => "__vybe_str_remove_range",
        "pascal.str_insert" => "__vybe_str_insert",
        "pascal.sort_in_place" => "__vybe_sort_in_place",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
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
