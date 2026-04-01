use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_compiler_common::expressions;

#[test]
fn ternary_emits_jumps() {
    let mut chunk = Chunk::new("test");
    // Simulate: condition already on stack
    chunk.emit_op(Op::r#true, 0); // condition
    let false_jump = expressions::emit_ternary_start(&mut chunk, 0);
    chunk.emit_op(Op::i32_const_1, 0); // then value
    let end_jump = expressions::emit_ternary_middle(&mut chunk, false_jump, 0);
    chunk.emit_op(Op::i32_const_0, 0); // else value
    expressions::emit_ternary_end(&mut chunk, end_jump);
    assert!(chunk.code.len() > 5, "ternary should emit jump structure");
}

#[test]
fn and_short_circuit() {
    let mut chunk = Chunk::new("test");
    chunk.emit_op(Op::r#true, 0); // left
    let jump = expressions::emit_and_start(&mut chunk, 0);
    chunk.emit_op(Op::r#false, 0); // right
    expressions::emit_short_circuit_end(&mut chunk, jump);
    assert!(chunk.code.len() > 3);
}

#[test]
fn or_short_circuit() {
    let mut chunk = Chunk::new("test");
    chunk.emit_op(Op::r#false, 0); // left
    let jump = expressions::emit_or_start(&mut chunk, 0);
    chunk.emit_op(Op::r#true, 0); // right
    expressions::emit_short_circuit_end(&mut chunk, jump);
    assert!(chunk.code.len() > 3);
}

#[test]
fn null_coalesce() {
    let mut chunk = Chunk::new("test");
    chunk.emit_op(Op::null, 0); // left
    let (_null_jump, end_jump) = expressions::emit_null_coalesce_start(&mut chunk, 0);
    chunk.emit_op(Op::i32_const_1, 0); // right (default)
    expressions::emit_null_coalesce_end(&mut chunk, end_jump);
    assert!(chunk.code.len() > 4);
}

#[test]
fn null_safe_access() {
    let mut chunk = Chunk::new("test");
    chunk.emit_op(Op::null, 0); // object
    let (skip, _) = expressions::emit_null_safe_start(&mut chunk, 0);
    // member access would go here
    chunk.emit_op(Op::drop, 0); // placeholder
    expressions::emit_null_safe_end(&mut chunk, skip, 0);
    assert!(chunk.code.len() > 3);
}
