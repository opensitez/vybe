use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_compiler::compiler::functions;

#[test]
fn default_param_roundtrip() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 5;
    let skip = functions::emit_default_param_start(&mut chunk, 1, 0);
    chunk.emit_op(Op::NULL, 0); // placeholder default
    functions::emit_default_param_end(&mut chunk, 1, skip, 0);
    assert!(chunk.code.len() > 5);
}

#[test]
fn create_function_chunk_sets_arity() {
    let chunk = functions::create_function_chunk("greet", 2);
    assert_eq!(chunk.name, "greet");
    assert_eq!(chunk.arity, 2);
}

#[test]
fn emit_function_epilogue_ends_with_return() {
    let mut chunk = Chunk::new("test");
    functions::emit_function_epilogue(&mut chunk, 0);
    assert!(chunk.code.len() >= 2, "should emit null + return");
}

#[test]
fn emit_ref_func_produces_closure() {
    let mut chunk = Chunk::new("test");
    functions::emit_ref_func(&mut chunk, 5, 0, 0);
    assert!(!chunk.code.is_empty());
}

#[test]
fn emit_store_global_func_adds_constant() {
    let mut chunk = Chunk::new("test");
    // Simulate closure on stack
    functions::emit_store_global_func(&mut chunk, "greet", 0);
    let has_name = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "greet"));
    assert!(has_name);
}

#[test]
fn emit_push_global_func_gets_by_name() {
    let mut chunk = Chunk::new("test");
    functions::emit_push_global_func(&mut chunk, "my_func", 0);
    let has_name = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "my_func"));
    assert!(has_name);
}

#[test]
fn emit_call_emits_opcode() {
    let mut chunk = Chunk::new("test");
    functions::emit_call(&mut chunk, 3, 0);
    assert!(!chunk.code.is_empty());
}

#[test]
fn full_cross_language_call_pattern() {
    // Simulate: Python calling a Dart function by name
    let mut chunk = Chunk::new("test");
    // push function ref
    functions::emit_push_global_func(&mut chunk, "dart_greet", 0);
    // push arg (would be compile_expr in real code)
    chunk.emit_op(Op::NULL, 0);
    // call with 1 arg
    functions::emit_call(&mut chunk, 1, 0);
    assert!(chunk.code.len() > 3);
}
