use vybe_compiler::primitives::instructions::host;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.json_parse" => {
            if argc >= 2 {
                chunks[current].emit_op(Op::DROP, line);
            }
            emit_json_text_coerce(&mut chunks[current], line);
            vybe_compiler::primitives::json::emit_parse_or_null(chunks, current, line);
        }
        _ => return false,
    }
    true
}

fn emit_json_text_coerce(chunk: &mut Chunk, line: u32) {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "wasm:js-string", "test", 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_else(line);
    host::emit(chunk, "web:encoding", "decoderNew", 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "ecma:uint8array", "newFromIterable", 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_end(line);
    host::emit(chunk, "web:encoding", "decode", 2, line);
    chunk.emit_end(line);
}
