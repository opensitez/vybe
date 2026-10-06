// Isolated compiler-only candidate; no production installation.
use vybe_runtime::{Chunk, Op};
#[path = "text_optimization_htmlentities_frozen_named.rs"]
mod named;

pub fn emit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_language_php::emitter::string_adapter::emit_htmlspecialchars(chunks, current, argc, line);
    let chunk = &chunks[current];
    let missing = [("ecma:regexp", "test"), ("wasm:js-boolean", "cast")]
        .into_iter()
        .filter(|(module, name)| {
            !chunk
                .imports
                .iter()
                .any(|i| i.module == *module && i.name == *name)
        })
        .count();
    // Initial coercion/basic passes are identical; select the legacy named chain
    // before adding guard imports, code, or scratch if the old frame is full.
    if chunk.local_count == u16::MAX || chunk.imports.len() + missing > usize::from(u16::MAX) + 1 {
        named::emit_named(&mut chunks[current], line);
        return;
    }
    emit_named_guard(&mut chunks[current], line);
}

pub fn emit_named_guard(chunk: &mut Chunk, line: u32) {
    let string = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, string, line);
    chunk.emit_op_u16(Op::LOCAL_GET, string, line);
    chunk.emit_string_const("[^\\x00-\\x7f]", line);
    chunk.emit_op_u16(Op::LOCAL_GET, string, line);
    let test = chunk.add_import("ecma:regexp", "test");
    chunk.emit_call(test, 2, line);
    // Keep the result String as an explicit native block operand, then release
    // the scratch reference before the old74-pass chain. The pure regex precedes
    // release; its checked Bool cast remains adjacent to IF afterward.
    vybe_compiler::primitives::instructions::core_wasm::null(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, string, line);
    // Keep the checked i32 producer adjacent to its condition consumer.
    let cast = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(cast, 1, line);
    chunk.emit_if_params(line, 1, 1);
    named::emit_named(chunk, line);
    chunk.emit_else(line);
    chunk.emit_end(line);
}
