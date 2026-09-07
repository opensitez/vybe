//! Go runtime-surface helpers routed via `common:go.*`.

use vybe_runtime::Chunk;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    let _ = argc;
    match name {
        "go.panic" => vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line),
        _ => return false,
    }
    true
}
