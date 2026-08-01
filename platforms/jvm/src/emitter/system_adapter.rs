use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(value, line);
}

fn call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

/// `java.lang.System.getProperty(key[, default])`.
pub fn emit_get_property(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let default_slot = chunk.alloc_scratch(1);
    let key_slot = chunk.alloc_scratch(1);

    if argc >= 2 {
        chunk.emit_op_u16(Op::LOCAL_SET, default_slot, line);
    } else {
        chunk.emit_op(Op::NULL, line);
        chunk.emit_op_u16(Op::LOCAL_SET, default_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    push_str(chunk, "java.io.tmpdir", line);
    call(chunk, "wasm:js-string", "equals", 2, line);
    chunk.emit_if_value(line);
    push_str(chunk, "/tmp", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    push_str(chunk, "file.separator", line);
    call(chunk, "wasm:js-string", "equals", 2, line);
    chunk.emit_if_value(line);
    push_str(chunk, "/", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    push_str(chunk, "line.separator", line);
    call(chunk, "wasm:js-string", "equals", 2, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("\n", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, default_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}
