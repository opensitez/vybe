use vybe_bytecode::Chunk;

use super::host;

pub fn is_object(c: &mut Chunk, line: u32) {
    host::emit(c, "ecma:value", "typeof", 1, line);
    c.emit_string_const("object", line);
    host::emit(c, "wasm:js-string", "equals", 2, line);
}

pub fn is_func(c: &mut Chunk, line: u32) {
    host::emit(c, "ecma:value", "typeof", 1, line);
    c.emit_string_const("function", line);
    host::emit(c, "wasm:js-string", "equals", 2, line);
}

pub fn string_reverse(c: &mut Chunk, line: u32) {
    c.emit_string_const("", line);
    host::emit(c, "ecma:string", "split", 2, line);
    host::emit(c, "ecma:array", "reverse", 1, line);
    c.emit_string_const("", line);
    host::emit(c, "ecma:array", "join", 2, line);
}
