use vybe_bytecode::Chunk;
use vybe_compiler_common::functions;

#[test]
fn emit_default_param_roundtrip() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 5;
    let skip = functions::emit_default_param_start(&mut chunk, 1, 0);
    // Caller would compile default expression here — push a constant as placeholder
    chunk.emit_op(vybe_bytecode::opcode::Op::null, 0);
    functions::emit_default_param_end(&mut chunk, 1, skip, 0);
    assert!(chunk.code.len() > 5, "default param check should emit opcodes");
}
