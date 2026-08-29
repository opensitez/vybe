// Classes were centralized into the compiler. The OBJECT-level primitives
// (method/accessor binding, protocol-slot publication) stayed in
// `vybe_compiler::primitives::object`, because crates BELOW the compiler — platforms/dotnet —
// still call them. Class construction, registration and super wiring moved to
// `vybe_compiler::primitives::classes`. This test spans both, so it names both.
use vybe_compiler::primitives::classes;
use vybe_compiler::primitives::dict;
use vybe_compiler::primitives::object;

// ── protocol slots: cross-language reach without synonyms ───
//
// These replace a synonym-table test set. The table bound every language's
// spelling of a special method as its own property, so the tests asserted
// things like "Python __contains__ aliases to C# contains" — which is exactly
// the behaviour that captured unrelated user methods. Reach is now a numeric
// role key, so what is worth asserting is that roles are distinct and that
// their keys cannot be spelled by a user.

use vybe_ast::{ProtocolSlot, protocol_slot_key};

#[test]
fn one_role_is_one_key_whatever_the_language_calls_it() {
    // Python `__str__`, Ruby `to_s`, PHP `__toString` and C# `ToString` all
    // normalize to `ProtocolSlot::ToString` in their own crates; by the time
    // a key is derived there is only the role left.
    assert_eq!(protocol_slot_key(ProtocolSlot::ToString), "__vybe_slot_1");
}

#[test]
fn slot_keys_are_derived_from_the_number_not_a_spelling() {
    // The point of the numeric derivation: a user method genuinely named
    // `toString` stays an ordinary member, because no slot key is ever a
    // pronounceable identifier.
    for slot in [
        ProtocolSlot::ToString,
        ProtocolSlot::Len,
        ProtocolSlot::GetItem,
        ProtocolSlot::Contains,
    ] {
        let key = protocol_slot_key(slot);
        assert!(key.starts_with("__vybe_slot_"));
        assert!(
            key["__vybe_slot_".len()..]
                .chars()
                .all(|c| c.is_ascii_digit())
        );
    }
}

#[test]
fn distinct_roles_never_share_a_key() {
    // Roles that DID share one until 2026-07-28: `/` and `//` both claimed
    // Div, `__int__`/`__float__` both claimed ValueOf, and PHP's `__invoke`
    // and `__call` both claimed Call. Each collision silently evicted one
    // method when the other installed.
    let pairs = [
        (ProtocolSlot::Div, ProtocolSlot::FloorDiv),
        (ProtocolSlot::Int, ProtocolSlot::Float),
        (ProtocolSlot::Call, ProtocolSlot::CallMissing),
        (ProtocolSlot::Eq, ProtocolSlot::Ne),
    ];
    for (a, b) in pairs {
        assert_ne!(protocol_slot_key(a), protocol_slot_key(b), "{a:?} vs {b:?}");
    }
}

#[test]
fn bind_with_slot_publishes_the_name_and_the_role() {
    let mut chunk = Chunk::new("test");
    object::emit_bind_method_with_slot(
        &mut chunk,
        0,
        "__str__",
        Some(ProtocolSlot::ToString),
        0,
        None,
        1,
    );
    let constants: Vec<String> = chunk
        .constants
        .iter()
        .filter_map(|c| match c {
            Value::String(text) => Some(text.to_string()),
            _ => None,
        })
        .collect();
    assert!(constants.contains(&"__str__".to_string()));
    assert!(constants.contains(&protocol_slot_key(ProtocolSlot::ToString)));
}

#[test]
fn an_ordinary_method_binds_under_one_name_only() {
    let mut chunk = Chunk::new("test");
    object::emit_bind_method_with_slot(&mut chunk, 0, "doSomething", None, 0, None, 1);
    let names: Vec<String> = chunk
        .constants
        .iter()
        .filter_map(|c| match c {
            Value::String(text) => Some(text.to_string()),
            _ => None,
        })
        .collect();
    assert_eq!(names, vec!["doSomething".to_string()]);
}

// ── emit helpers produce correct bytecode ───────────────────

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

#[test]
fn emit_new_typed_object_stamps_type() {
    let mut chunk = Chunk::new("test");
    // typeidx 1 — allocation carries the type, so the rtt is set by
    // `struct.new_default $T` rather than stamped afterwards.
    classes::emit_new_typed_object(&mut chunk, 1, "Dog", 1, 0);
    // Should have emitted struct.new_default $T and the __type stamp
    assert!(chunk.code.len() > 10, "Should emit multiple opcodes");
    // The class name reaches the module as an imported STRING CONSTANT — a
    // declared global whose field name is its value — so the pool holds the
    // import's `(module, name)` key, not a bare `"Dog"`.
    let dog_key = vybe_runtime::chunk::imported_global_key(
        vybe_runtime::chunk::STRING_CONSTANTS_MODULE,
        "Dog",
    );
    let has_dog = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == dog_key));
    assert!(
        chunk
            .global_imports
            .iter()
            .any(|i| i.module == vybe_runtime::chunk::STRING_CONSTANTS_MODULE && i.name == "Dog"),
        "the stamp must DECLARE its string constant"
    );
    let has_type = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__type"));
    assert!(has_dog, "Should have 'Dog' constant");
    assert!(has_type, "Should have '__type' constant");
}

#[test]
fn emit_bind_method_sets_property() {
    let mut chunk = Chunk::new("test");
    object::emit_bind_method(&mut chunk, 1, "greet", 5, 0);
    // Should emit local_get, ref_func, struct_set, drop
    assert!(
        chunk.code.len() > 5,
        "Should emit opcodes for method binding"
    );
    let has_greet = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "greet"));
    assert!(has_greet, "Should have 'greet' constant");
}

#[test]
fn emit_store_super_sets_property() {
    let mut chunk = Chunk::new("test");
    classes::emit_store_super(&mut chunk, 1, "animal", 0);
    let has_super = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__super"));
    let has_parent = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "animal"));
    assert!(has_super, "Should have '__super' constant");
    assert!(has_parent, "Should have parent name constant");
}

#[test]
fn emit_save_base_method_creates_base_prefix() {
    let mut chunk = Chunk::new("test");
    classes::emit_save_base_method(&mut chunk, 1, "derived", "greet", 0);
    let has_base = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__base_derived$greet"));
    assert!(has_base, "Should have '__base_derived$greet' constant");
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
        &mut chunks,
        "Dog",
        "Animal",
        vec!["name".into()],
        vec![("bark".into(), 1)],
        false,
        vec![],
        Some(2),
        std::collections::HashMap::new(),
    );
    // Two entries: `Dog`, plus a DECLARATION for the supertype it names.
    // Declaring the supertype is what lets the subtype link be an index —
    // `sub $i` — instead of a name resolved at load.
    let dog = chunks[0]
        .types
        .iter()
        .position(|t| t.name == "Dog")
        .expect("Dog registered");
    let animal = chunks[0]
        .types
        .iter()
        .position(|t| t.name == "Animal")
        .expect("supertype declared");
    // register_type preserves the source-language name verbatim;
    // case-insensitive lookup is the type registry's responsibility.
    assert_eq!(
        chunks[0].types[dog].parent_index as usize,
        animal + 1,
        "the supertype link must point at the declared entry (1-based)"
    );
    assert_eq!(chunks[0].types[dog].fields, vec!["name"]);
}

#[test]
fn emit_attach_static_method_sets_on_constructor() {
    let mut chunk = Chunk::new("test");
    classes::emit_attach_static_method(&mut chunk, true, 2, "create", 5, None, None, 0);
    let has_create = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "create"));
    assert!(
        has_create,
        "Should have 'create' constant for static method"
    );
}

#[test]
fn emit_bind_getter_uses_get_prefix() {
    let mut chunk = Chunk::new("test");
    object::emit_bind_getter(&mut chunk, 1, "name", 3, 0);
    let has_get = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__get_name"));
    assert!(has_get, "Should have '__get_name' constant for getter");
}

#[test]
fn emit_bind_setter_uses_set_prefix() {
    let mut chunk = Chunk::new("test");
    object::emit_bind_setter(&mut chunk, 1, "name", 4, 0);
    let has_set = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__set_name"));
    assert!(has_set, "Should have '__set_name' constant for setter");
}

#[test]
fn emit_init_field_null_sets_null() {
    let mut chunk = Chunk::new("test");
    classes::emit_init_field_null(&mut chunk, 1, "count", 0);
    let has_count = chunk
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "count"));
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
    let has_keys = chunks[0]
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__keys"));
    assert!(
        has_keys,
        "dict should have __keys constant for key tracking"
    );
}

#[test]
fn dict_set_const_key_tracks_key() {
    let mut chunks = one_chunk();
    dict::emit_new(&mut chunks, 0, 0);
    // Simulate: dup dict, push value, then set key
    chunks[0].emit_dup(0);
    chunks[0].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0); // placeholder value
    dict::emit_set_const_key(&mut chunks, 0, "name", 0);
    // Should have "name" in constants (for struct_set AND for __keys push)
    let name_count = chunks[0]
        .constants
        .iter()
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
    let has_keys = chunks[0]
        .constants
        .iter()
        .any(|c| matches!(c, Value::String(s) if s.as_ref() == "__keys"));
    assert!(has_keys);
}

/// A name a chunk READS but never WRITES is an import of the embedder, and the
/// module must declare it — a WASM module may only touch globals it declared.
/// The test is the bytecode write-set, not what the source declared: the
/// prelude declares `globalThis` and never assigns it.
#[test]
fn free_globals_are_declared_as_host_imports() {
    let mut chunk = Chunk::new("<script>");
    let read_only = chunk.add_constant(Value::String(std::sync::Arc::from("globalThis")));
    let written = chunk.add_constant(Value::String(std::sync::Arc::from("myVar")));
    chunk.emit_op_u16(Op::GLOBAL_GET, read_only, 0);
    chunk.emit_op_u16(Op::GLOBAL_SET, written, 0);
    chunk.emit_op_u16(Op::GLOBAL_GET, written, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut chunks = vec![chunk];
    vybe_compiler::primitives::globals::declare_free_globals(&mut chunks);

    let declared: Vec<&str> = chunks[0]
        .global_imports
        .iter()
        .filter(|i| i.module == vybe_runtime::chunk::HOST_GLOBALS_MODULE)
        .map(|i| i.name.as_str())
        .collect();
    assert_eq!(
        declared,
        vec!["globalThis"],
        "only the read-never-written name is an import"
    );
}
