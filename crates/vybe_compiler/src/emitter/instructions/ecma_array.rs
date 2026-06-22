use vybe_bytecode::Chunk;

pub fn is_array(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:array", "isArray");
    c.emit_call(idx, 1, line);
}
