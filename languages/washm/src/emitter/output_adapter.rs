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
        "washm.echo" => emit_echo(chunks, current, argc, line),
        "washm.printf" => emit_printf(chunks, current, argc, line),
        "washm.sprintf" => {
            vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line)
        }
        "washm.sprintf_array" => {
            vybe_compiler::primitives::sprintf::emit_sprintf_from_array(chunks, current, line)
        }
        _ => return false,
    }
    true
}

fn emit_echo(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("\n", line);
        vybe_compiler::primitives::io::emit_write_or_buffer(chunks, current, line);
        chunks[current].emit_i32_const(0, line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }

    emit_stringified_slot(chunks, current, base, line);
    for offset in 1..argc as u16 {
        chunks[current].emit_string_const(" ", line);
        host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
        emit_stringified_slot(chunks, current, base + offset, line);
        host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
    }
    chunks[current].emit_string_const("\n", line);
    host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
    vybe_compiler::primitives::io::emit_write_or_buffer(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
}

fn emit_printf(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
    vybe_compiler::primitives::io::emit_write_or_buffer(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
}

fn emit_stringified_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    vybe_compiler::primitives::expressions::emit_rich_to_string(&mut chunks[current], slot, line);
}
