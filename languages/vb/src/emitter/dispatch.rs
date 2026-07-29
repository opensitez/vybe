//! Auto-extracted `vb.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    let _ = (name, chunks, current, argc, line);
    false
}
