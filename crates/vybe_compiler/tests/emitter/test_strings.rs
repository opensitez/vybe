use vybe_compiler::primitives::strings;
use vybe_runtime::{Chunk, Value};

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
    // Zero parts is the empty string — declared like any other string
    // constant. The spec allows an empty import field name, and the JS-API
    // test vector for string constants includes `''` explicitly.
    assert!(
        chunk
            .global_imports
            .iter()
            .any(|i| i.module == vybe_runtime::chunk::STRING_CONSTANTS_MODULE
                && i.name.is_empty()),
        "0 parts should push the empty string"
    );
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
    // A string literal is an imported global (js-string-builtins § String
    // constants): the pool holds the import key, and the import itself is
    // declared on the chunk.
    let key = vybe_runtime::chunk::imported_global_key(
        vybe_runtime::chunk::STRING_CONSTANTS_MODULE,
        "hello",
    );
    let has_hello = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == key));
    assert!(has_hello, "should reference the 'hello' string constant");
    assert!(
        chunk
            .global_imports
            .iter()
            .any(|i| i.module == vybe_runtime::chunk::STRING_CONSTANTS_MODULE
                && i.name == "hello"),
        "should declare the 'hello' string constant as an imported global"
    );
}
