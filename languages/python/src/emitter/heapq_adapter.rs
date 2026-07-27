//! Python `heapq` adapter.
//!
//! The heap algorithm lives in `vybe_compiler::compiler::heap` so Python, Go, Java, and
//! C# priority-queue surfaces can converge on the same bytecode behavior.

use vybe_bytecode::Chunk;

pub fn emit_heapify(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::compiler::heap::emit_heapify(chunks, current, argc, line);
}

pub fn emit_heappush(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::compiler::heap::emit_push(chunks, current, argc, line);
}

pub fn emit_heappop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::compiler::heap::emit_pop(chunks, current, argc, line);
}

pub fn emit_heapreplace(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::compiler::heap::emit_replace(chunks, current, argc, line);
}

pub fn emit_heappushpop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::compiler::heap::emit_push_pop(chunks, current, argc, line);
}

pub fn emit_nsmallest(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::compiler::heap::emit_nsmallest(chunks, current, argc, line);
}

pub fn emit_nlargest(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::compiler::heap::emit_nlargest(chunks, current, argc, line);
}

pub fn emit_merge(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::compiler::heap::emit_merge(chunks, current, argc, line);
}
