//! Python runtime-surface emitters.
//!
//! These are routed from the Python profile through `common:python.*`.
//! Keep Python-specific call shapes here instead of sending them through
//! the old runtime-helper function table.

use crate::emitter::{collections, target::Target};
use vybe_bytecode::Chunk;

/// Python `range(...)`.
///
/// The common one-argument form is emitted inline as a WASM loop. The
/// multi-argument forms still fall back to the shared runtime helper for
/// now because they need Python's nullable-argument reshaping semantics.
pub fn emit_range(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    collections::emit_range_targeted(chunks, current, argc, &Target::wasm(), line);
}

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    let global = match name {
        "python.hex" => "__vybe_pyhex",
        "python.oct" => "__vybe_pyoct",
        "python.bin" => "__vybe_pybin",
        "python.bytes" | "python.encode" => "__vybe_to_bytes",
        "python.enumerate" => "__vybe_enumerate",
        "python.zip" => "__vybe_zip",
        "python.map" => "__vybe_pymap",
        "python.filter" => "__vybe_pyfilter",
        "python.any" => "__vybe_pyany",
        "python.all" => "__vybe_pyall",
        "python.iter" => "__vybe_pyiter",
        "python.next" => "__vybe_pynext",
        "python.isinf" => "__vybe_isinf",
        "python.random_choice" => "__vybe_rand_choice",
        "python.random_shuffle" => "__vybe_rand_shuffle",
        "python.random_sample" => "__vybe_rand_sample",
        "python.instanceof" => "__vybe_instanceof",
        "python.callable" => "__vybe_callable",
        "python.id" => "__vybe_id",
        "python.hash" => "__vybe_hash",
        "python.regex_findall" => "__ecma_regexp_match_all_pat_first",
        "python.regex_sub" => "__ecma_regexp_replace_pat_first",
        "python.regex_split" => "__ecma_regexp_split_pat_first",
        "python.format_map" => "__vybe_format_map",
        "python.setdefault" => "__vybe_setdefault",
        "python.tostring" => "__vybe_tostring",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}
