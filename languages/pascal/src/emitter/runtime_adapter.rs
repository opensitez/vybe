//! Pascal runtime-surface helpers routed via `common:pascal.*`.

use vybe_emitter::collections;
use vybe_bytecode::Chunk;
use vybe_bytecode::Op;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    if name == "pascal.tostring" {
        let to_str = chunks[0].add_import("ecma:string", "String");
        chunks[current].emit_op_u16(Op::CALL_IMPORT, to_str, line);
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
