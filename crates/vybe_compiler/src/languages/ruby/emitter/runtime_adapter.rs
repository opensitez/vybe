//! Ruby runtime-surface emitters routed via `common:ruby.*`.

use crate::emitter::collections;
use vybe_bytecode::Chunk;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    let global = match name {
        "ruby.tostring" => "__vybe_tostring",
        "ruby.bytes" => "__vybe_to_bytes",
        "ruby.instanceof" => "__vybe_instanceof",
        "ruby.hash" => "__vybe_hash",
        "ruby.id" => "__vybe_id",
        "ruby.encoding" => "__vybe_encoding",
        "ruby.hex" => "__vybe_pyhex",
        "ruby.compact" => "__vybe_compact",
        "ruby.uniq" => "__vybe_uniq",
        "ruby.minmax" => "__vybe_minmax",
        "ruby.isempty" => "__vybe_isempty",
        "ruby.sample" => "__vybe_rand_choice",
        "ruby.shuffle" => "__vybe_rand_shuffle",
        "ruby.rotate" => "__vybe_rotate",
        "ruby.zip" => "__vybe_zip",
        "ruby.has_value" => "__vybe_has_value",
        "ruby.transform_values" => "__vybe_transform_values",
        "ruby.transform_keys" => "__vybe_transform_keys",
        "ruby.invert" => "__vybe_invert",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}
