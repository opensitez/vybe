use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::strings;

#[test]
fn emit_to_string_adds_import() {
    let mut chunk = Chunk::new("test");
    strings::emit_to_string(&mut chunk, 0);
    assert!(!chunk.code.is_empty(), "should emit host call");
}

#[test]
fn emit_concat_zero_parts() {
    let mut chunk = Chunk::new("test");
    strings::emit_concat(&mut chunk, 0, 0);
    // Should push empty string
    let has_empty = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == ""));
    assert!(has_empty, "0 parts should push empty string");
}

#[test]
fn emit_concat_one_part_noop() {
    let mut chunk = Chunk::new("test");
    let before = chunk.code.len();
    strings::emit_concat(&mut chunk, 1, 0);
    assert_eq!(chunk.code.len(), before, "1 part should be no-op");
}

#[test]
fn emit_concat_multiple_parts() {
    let mut chunk = Chunk::new("test");
    strings::emit_concat(&mut chunk, 5, 0);
    assert!(!chunk.code.is_empty(), "5 parts should emit concat opcodes");
}

#[test]
fn emit_literal_part_adds_constant() {
    let mut chunk = Chunk::new("test");
    strings::emit_literal_part(&mut chunk, "hello", 0);
    let has_hello = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "hello"));
    assert!(has_hello, "should add 'hello' constant");
}
