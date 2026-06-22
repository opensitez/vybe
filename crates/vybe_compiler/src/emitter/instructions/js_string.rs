use vybe_bytecode::Chunk;

pub fn concat(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-string", "concat");
    c.emit_call(idx, 2, line);
}

pub fn length(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-string", "length");
    c.emit_call(idx, 1, line);
}

pub fn substring(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-string", "substring");
    c.emit_call(idx, 3, line);
}

pub fn test(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-string", "test");
    c.emit_call(idx, 1, line);
}

pub fn equals(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-string", "equals");
    c.emit_call(idx, 2, line);
}

pub fn compare(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-string", "compare");
    c.emit_call(idx, 2, line);
}

pub fn char_code_at(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-string", "charCodeAt");
    c.emit_call(idx, 2, line);
}

pub fn from_char_code(c: &mut Chunk, line: u32) {
    let idx = c.add_import("wasm:js-string", "fromCharCode");
    c.emit_call(idx, 1, line);
}
