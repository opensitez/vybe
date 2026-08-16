use vybe_compiler::primitives::expressions;
use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

#[test]
fn ternary_emits_jumps() {
    let mut chunk = Chunk::new("test");
    // Simulate: condition already on stack
    chunk.emit_i32_const(1, 0); // condition
    let false_jump = expressions::emit_ternary_start(&mut chunk, 0);
    chunk.emit_i32_const(1, 0); // then value
    let end_jump = expressions::emit_ternary_middle(&mut chunk, false_jump, 0);
    chunk.emit_i32_const(0, 0); // else value
    expressions::emit_ternary_end(&mut chunk, end_jump);
    assert!(chunk.code.len() > 5, "ternary should emit jump structure");
}

#[test]
fn and_short_circuit() {
    let mut chunk = Chunk::new("test");
    chunk.emit_i32_const(1, 0); // left
    let jump = expressions::emit_and_start(&mut chunk, 0);
    chunk.emit_i32_const(0, 0); // right
    expressions::emit_short_circuit_end(&mut chunk, jump);
    assert!(chunk.code.len() > 3);
}

#[test]
fn or_short_circuit() {
    let mut chunk = Chunk::new("test");
    chunk.emit_i32_const(0, 0); // left
    let jump = expressions::emit_or_start(&mut chunk, 0);
    chunk.emit_i32_const(1, 0); // right
    expressions::emit_short_circuit_end(&mut chunk, jump);
    assert!(chunk.code.len() > 3);
}

#[test]
fn null_coalesce() {
    let mut chunk = Chunk::new("test");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0); // left
    let (_null_jump, end_jump) = expressions::emit_null_coalesce_start(&mut chunk, 0);
    chunk.emit_i32_const(1, 0); // right (default)
    expressions::emit_null_coalesce_end(&mut chunk, end_jump);
    assert!(chunk.code.len() > 4);
}

#[test]
fn null_safe_access() {
    let mut chunk = Chunk::new("test");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0); // object
    let (skip, _) = expressions::emit_null_safe_start(&mut chunk, 0);
    // member access would go here
    chunk.emit_op(Op::DROP, 0); // placeholder
    expressions::emit_null_safe_end(&mut chunk, skip, 0);
    assert!(chunk.code.len() > 3);
}

#[test]
fn rich_compare_locals_emits_dispatch() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 5;
    // Takes the whole chunk vector + the current index: the fallback emitter
    // needs the vector, so a single `&mut Chunk` cannot be held across it.
    let mut chunks = vec![chunk];
    expressions::emit_rich_compare_locals(
        &mut chunks,
        0,
        1,
        2,
        "__lt__",
        expressions::RichFallback::Op(ops::emit_dyn_lt),
        0,
        // This test asserts the DISPATCH chain is emitted, so the operand must
        // not be a known number — that is the case the chain is skipped for.
        false,
    );
    let chunk = chunks.remove(0);
    // Should have: struct_get, dup, ref_is_null, br_if_true, call_ref, br, drop, drop, dyn_lt
    assert!(
        chunk.code.len() > 15,
        "rich compare should emit dispatch bytecode"
    );
    let has_lt = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__lt__"));
    assert!(has_lt, "should have '__lt__' constant for struct_get");
}

#[test]
fn smart_length_emits_dispatch() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 5;
    expressions::emit_smart_length(&mut chunk, 1, 0);
    assert!(
        chunk.code.len() > 15,
        "smart length should emit dispatch bytecode"
    );
    let has_len = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__get_length"));
    assert!(
        has_len,
        "should have '__get_length' constant for getter check"
    );
}

#[test]
fn rich_compare_fallback_emits_primitive() {
    // When no dunder found, should fall back to the primitive op
    let mut chunk = Chunk::new("test");
    chunk.local_count = 5;
    expressions::emit_rich_compare(&mut chunk, "__lt__", ops::emit_dyn_lt, 0);
    // Simple fallback version just emits the opcode directly
    assert!(!chunk.code.is_empty());
}
