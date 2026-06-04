//! PHP runtime-surface helpers routed via `common:php.*`.

use crate::emitter::collections;
use vybe_bytecode::Chunk;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    let global = match name {
        "php.regex_match_all_pat_first" => "__ecma_regexp_match_all_pat_first",
        "php.regex_replace_pat_first" => "__ecma_regexp_replace_pat_first",
        "php.isnumeric" => "__vybe_isnumeric",
        "php.zip" => "__vybe_zip",
        "php.sort_in_place" => "__vybe_sort_in_place",
        "php.sort_with_comparator" => "__vybe_sort_with_comparator",
        "php.uniq" => "__vybe_uniq",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}
