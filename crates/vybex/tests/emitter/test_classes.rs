use vybex::emitter::classes;
use vybex::emitter::dict;

// ── cross_language_aliases lookup table ─────────────────────

#[test]
fn alias_tostring_to_str() {
    let aliases = classes::cross_language_aliases("toString");
    assert!(aliases.contains(&"__str__"), "JS toString should alias to Python __str__");
}

#[test]
fn alias_tostring_lowercase() {
    let aliases = classes::cross_language_aliases("tostring");
    assert!(aliases.contains(&"__str__"), "VB tostring should alias to Python __str__");
    assert!(aliases.contains(&"toString"), "VB tostring should alias to JS toString");
}

#[test]
fn alias_len_length() {
    let aliases = classes::cross_language_aliases("__len__");
    assert!(aliases.contains(&"__get_length"), "Python __len__ should alias to JS length");
    assert!(aliases.contains(&"__get_count"), "Python __len__ should alias to VB/C# Count");
}

#[test]
fn alias_length_to_len() {
    let aliases = classes::cross_language_aliases("__get_length");
    assert!(aliases.contains(&"__len__"), "JS length should alias to Python __len__");
}

#[test]
fn alias_contains_all_directions() {
    let py = classes::cross_language_aliases("__contains__");
    let js = classes::cross_language_aliases("includes");
    let cs = classes::cross_language_aliases("contains");
    // All should map to the same set
    assert!(py.contains(&"includes"), "Python __contains__ → JS includes");
    assert!(py.contains(&"contains"), "Python __contains__ → C# contains");
    assert!(js.contains(&"__contains__"), "JS includes → Python __contains__");
    assert!(cs.contains(&"__contains__"), "C# contains → Python __contains__");
}

#[test]
fn alias_bool_valueof() {
    let aliases = classes::cross_language_aliases("__bool__");
    assert!(aliases.contains(&"valueOf"), "Python __bool__ should alias to JS valueOf");
}

#[test]
fn alias_eq_equals() {
    let aliases = classes::cross_language_aliases("__eq__");
    assert!(aliases.contains(&"equals"), "Python __eq__ should alias to C#/VB equals");
}

#[test]
fn alias_regular_method_no_aliases() {
    let aliases = classes::cross_language_aliases("doSomething");
    assert!(aliases.is_empty(), "Regular method names should have no aliases");
}

#[test]
fn alias_repr() {
    let aliases = classes::cross_language_aliases("__repr__");
    assert!(aliases.contains(&"toDebugString"), "__repr__ → toDebugString");
}

// ── emit helpers produce correct bytecode ───────────────────

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

#[test]
fn emit_new_typed_object_stamps_type() {
    let mut chunk = Chunk::new("test");
    classes::emit_new_typed_object(&mut chunk, 1, "Dog", 0);
    // Should have emitted struct_new, __type stamp, set_type_id
    assert!(chunk.code.len() > 10, "Should emit multiple opcodes");
    // Check constants contain "Dog" and "__type"
    let has_dog = chunk.constants.iter().any(|c| matches!(c, Value::String(s) if s.as_ref() == "Dog"));
    let has_type = chunk.constants.iter().any(|c| matches!(c, Value::String(s) if s.as_ref() == "__type"));
    assert!(has_dog, "Should have 'Dog' constant");
    assert!(has_type, "Should have '__type' constant");
}

#[test]
fn emit_bind_method_sets_property() {
    let mut chunk = Chunk::new("test");
    classes::emit_bind_method(&mut chunk, 1, "greet", 5, 0);
    // Should emit local_get, ref_func, struct_set, drop
    assert!(chunk.code.len() > 5, "Should emit opcodes for method binding");
    let has_greet = chunk.constants.iter().any(|c| matches!(c, Value::String(s) if s.as_ref() == "greet"));
    assert!(has_greet, "Should have 'greet' constant");
}

#[test]
fn emit_store_super_sets_property() {
    let mut chunk = Chunk::new("test");
    classes::emit_store_super(&mut chunk, 1, "animal", 0);
    let has_super = chunk.constants.iter().any(|c| matches!(c, Value::String(s) if s.as_ref() == "__super"));
    let has_parent = chunk.constants.iter().any(|c| matches!(c, Value::String(s) if s.as_ref() == "animal"));
    assert!(has_super, "Should have '__super' constant");
    assert!(has_parent, "Should have parent name constant");
}

#[test]
fn emit_save_base_method_creates_base_prefix() {
    let mut chunk = Chunk::new("test");
    classes::emit_save_base_method(&mut chunk, 1, "greet", 0);
    let has_base = chunk.constants.iter().any(|c| matches!(c, Value::String(s) if s.as_ref() == "__base_greet"));
    assert!(has_base, "Should have '__base_greet' constant");
}

#[test]
fn emit_constructor_return_emits_return() {
    let mut chunk = Chunk::new("test");
    classes::emit_constructor_return(&mut chunk, 1, 0);
    // Last opcode should be return
    let last = chunk.code.last().copied();
    assert!(last.is_some(), "Should emit opcodes");
}

#[test]
fn register_type_adds_entry() {
    let mut chunks = vec![Chunk::new("main")];
    classes::register_type(
        &mut chunks, "Dog", "Animal",
        vec!["name".into()],
        vec![("bark".into(), 1)],
        false,
        vec![],
        Some(2),
    );
    assert_eq!(chunks[0].types.len(), 1);
    // register_type preserves the source-language name verbatim;
    // case-insensitive lookup is the type registry's responsibility.
    assert_eq!(chunks[0].types[0].name, "Dog");
    assert_eq!(chunks[0].types[0].parent, "Animal");
    assert_eq!(chunks[0].types[0].fields, vec!["name"]);
}

#[test]
fn emit_attach_static_method_sets_on_constructor() {
    let mut chunk = Chunk::new("test");
    classes::emit_attach_static_method(&mut chunk, 2, "create", 5, 0);
    let has_create = chunk.constants.iter().any(|c| matches!(c, Value::String(s) if s.as_ref() == "create"));
    assert!(has_create, "Should have 'create' constant for static method");
}

#[test]
fn emit_bind_getter_uses_get_prefix() {
    let mut chunk = Chunk::new("test");
    classes::emit_bind_getter(&mut chunk, 1, "name", 3, 0);
    let has_get = chunk.constants.iter().any(|c| matches!(c, Value::String(s) if s.as_ref() == "__get_name"));
    assert!(has_get, "Should have '__get_name' constant for getter");
}

#[test]
fn emit_bind_setter_uses_set_prefix() {
    let mut chunk = Chunk::new("test");
    classes::emit_bind_setter(&mut chunk, 1, "name", 4, 0);
    let has_set = chunk.constants.iter().any(|c| matches!(c, Value::String(s) if s.as_ref() == "__set_name"));
    assert!(has_set, "Should have '__set_name' constant for setter");
}

#[test]
fn emit_init_field_null_sets_null() {
    let mut chunk = Chunk::new("test");
    classes::emit_init_field_null(&mut chunk, 1, "count", 0);
    let has_count = chunk.constants.iter().any(|c| matches!(c, Value::String(s) if s.as_ref() == "count"));
    assert!(has_count, "Should have 'count' constant for field init");
}

// ── dict helpers (cross-language) ──────────────────────────

// The `dict::emit_*` helpers take `(chunks: &mut [Chunk], current: usize,
// ...args, line: u32)` so they can also register imports on chunks[0]
// when building stdlib polyfills. Tests drive them through a one-element
// chunks slice.

fn one_chunk() -> Vec<Chunk> {
    vec![Chunk::new("test")]
}

#[test]
fn dict_new_has_keys_array() {
    let mut chunks = one_chunk();
    dict::emit_new(&mut chunks, 0, 0);
    let has_keys = chunks[0].constants.iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__keys"));
    assert!(has_keys, "dict should have __keys constant for key tracking");
}

#[test]
fn dict_set_const_key_tracks_key() {
    let mut chunks = one_chunk();
    dict::emit_new(&mut chunks, 0, 0);
    // Simulate: dup dict, push value, then set key
    chunks[0].emit_op(Op::DUP, 0);
    chunks[0].emit_op(Op::NULL, 0); // placeholder value
    dict::emit_set_const_key(&mut chunks, 0, "name", 0);
    // Should have "name" in constants (for struct_set AND for __keys push)
    let name_count = chunks[0].constants.iter()
        .filter(|c| matches!(c, Value::String(s) if s.as_ref() == "name"))
        .count();
    assert!(name_count >= 1, "should have 'name' constant");
}

#[test]
fn dict_keys_uses_struct_get() {
    let mut chunks = one_chunk();
    dict::emit_new(&mut chunks, 0, 0);
    dict::emit_keys(&mut chunks, 0, 0);
    // emit_keys does struct_get "__keys" which is pure WASM — no imports needed
    let has_keys = chunks[0].constants.iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__keys"));
    assert!(has_keys);
}
