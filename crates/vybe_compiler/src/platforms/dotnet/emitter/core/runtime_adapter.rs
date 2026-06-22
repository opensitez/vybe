//! Shared .NET-shaped helpers routed through runtime helper chunks.

use crate::emitter::collections;
use vybe_bytecode::Chunk;
use vybe_bytecode::Op;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    if name == "dotnet.tostring" {
        let to_str = chunks[0].add_import("ecma:string", "String");
        chunks[current].emit_op_u16(Op::CALL_IMPORT, to_str, line);
        chunks[current].emit(argc, line);
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
        "dotnet.sort_with_comparator" => "__vybe_sort_with_comparator",
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
