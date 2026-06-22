use vybe_bytecode::Chunk;

pub fn test(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-number", "test");
    c.emit_call(idx, 1, line);
}
