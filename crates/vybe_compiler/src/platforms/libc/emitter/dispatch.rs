//! libc `common:libc.*` emit dispatch.
//!
//! C stdio formatting is owned here under the libc platform, not by the
//! generic common formatter.

use vybe_bytecode::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "libc.stdio.printf" => {
            super::stdio_format::emit_sprintf(chunks, current, argc, line);
            let idx = chunks[current].add_import("wasi:logging/logging", "log");
            chunks[current].emit_call(idx, 1, line);
            true
        }
        "libc.stdio.sprintf" => {
            super::stdio_format::emit_sprintf(chunks, current, argc, line);
            true
        }
        "libc.stdio.vsprintf" => {
            super::stdio_format::emit_sprintf_from_array(chunks, current, line);
            true
        }
        _ => false,
    }
}
