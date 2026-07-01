use vybe_bytecode::Chunk;

use crate::emitter::{collections, strings};

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) -> bool {
    match name {
        "c.putchar" => {
            let idx = chunks[current].add_import("wasm:js-string", "fromCharCode");
            chunks[current].emit_call(idx, 1, line);
        }
        "c.strlen" => strings::emit_length(&mut chunks[current], line),
        "c.strupr" => strings::emit_to_upper(&mut chunks[current], line),
        "c.strlwr" => strings::emit_to_lower(&mut chunks[current], line),
        "c.strcmp" | "c.strncmp" | "c.memcmp" => {
            let idx = chunks[current].add_import("wasm:js-string", "compare");
            chunks[current].emit_call(idx, 2, line);
        }
        "c.atoi" | "c.atol" => {
            let idx = chunks[current].add_import("ecma:number", "parseInt");
            chunks[current].emit_call(idx, 1, line);
        }
        "c.qsort" => collections::emit_sort(chunks, current, line),
        _ => return false,
    }
    true
}
