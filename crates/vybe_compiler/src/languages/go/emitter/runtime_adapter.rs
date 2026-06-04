//! Go runtime-surface helpers routed via `common:go.*`.

use crate::emitter::collections;
use vybe_bytecode::Chunk;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    let global = match name {
        "go.regex_split_pat_first" => "__ecma_regexp_split_pat_first",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}
