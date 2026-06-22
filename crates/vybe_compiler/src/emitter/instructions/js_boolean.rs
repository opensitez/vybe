use vybe_bytecode::Chunk;

pub fn from_i32(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-boolean", "fromI32");
    c.emit_call(idx, 1, line);
}

pub fn test(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-boolean", "test");
    c.emit_call(idx, 1, line);
}
