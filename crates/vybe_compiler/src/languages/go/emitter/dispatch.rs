//! Go-specific common dispatch.

use vybe_bytecode::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    crate::languages::go::emitter::runtime_adapter::emit_helper(name, chunks, current, argc, line)
}
