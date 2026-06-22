use vybe_bytecode::Chunk;

pub fn typeof_op(c: &mut Chunk, line: u32) {
    let idx = c.add_import("ecma:value", "typeof");
    c.emit_call(idx, 1, line);
}

pub fn is_object(c: &mut Chunk, line: u32) {
    typeof_op(c, line);
    c.emit_string_const("object", line);
    let eq = c.add_import("wasm:js-string", "equals");
    c.emit_call(eq, 2, line);
}

pub fn is_func(c: &mut Chunk, line: u32) {
    typeof_op(c, line);
    c.emit_string_const("function", line);
    let eq = c.add_import("wasm:js-string", "equals");
    c.emit_call(eq, 2, line);
}
