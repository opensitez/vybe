use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

pub fn emit_oledb_connection_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    call_import(
        chunks,
        current,
        "wasi:sql/types",
        "connection.new",
        argc,
        line,
    );
}

pub fn emit_oledb_command_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    call_import(chunks, current, "wasi:sql/types", "command.new", argc, line);
}
