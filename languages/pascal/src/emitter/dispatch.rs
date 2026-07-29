//! Pascal-specific common dispatch.

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    crate::emitter::runtime_adapter::emit_helper(name, chunks, current, argc, line)
}
