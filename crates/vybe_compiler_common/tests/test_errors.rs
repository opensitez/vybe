use vybe_bytecode::Chunk;
use vybe_compiler_common::errors;

#[test]
fn emit_try_start_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    let catch_jump = errors::emit_try_start(&mut chunk, 0);
    assert!(chunk.code.len() > 2, "try_start should emit opcodes");
    assert!(catch_jump > 0, "should return a valid jump offset");
}

#[test]
fn emit_try_end_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    errors::emit_try_end(&mut chunk, 0);
    assert!(!chunk.code.is_empty(), "try_end should emit opcode");
}

#[test]
fn emit_throw_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    errors::emit_throw(&mut chunk, 0);
    assert!(!chunk.code.is_empty(), "throw should emit opcode");
}

#[test]
fn try_catch_roundtrip() {
    let mut chunk = Chunk::new("test");
    let catch_jump = errors::emit_try_start(&mut chunk, 0);
    // ... body would go here ...
    errors::emit_try_end(&mut chunk, 0);
    let skip = chunk.emit_jump(vybe_bytecode::opcode::Op::BR, 0);
    errors::patch_catch(&mut chunk, catch_jump);
    // ... handler would go here ...
    chunk.patch_jump(skip);
    assert!(chunk.code.len() > 5, "full try/catch should emit multiple opcodes");
}
