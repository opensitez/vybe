//! Tests for the WASM GC type system integration.
//!
//! Covers:
//! - TypeRegistry registration and get_id
//! - Subtype checking (is_subtype) with inheritance chains
//! - Method resolution through vtable (resolve_method)
//! - ref_test opcode with registered types
//! - Cross-language instanceof using __type and __types properties

use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Method, Op, TypeDef, VM, Value};

/// Declare `name` as this module's type 1 so a `ref.test` can carry an INDEX.
/// When the registry already knows the name — a host type, or one another
/// module registered — the declaration binds to it at load; that binding is
/// what makes a type reference an index on both sides.
fn declare_type(chunk: &mut Chunk, name: &str) {
    chunk.types.push(vybe_runtime::chunk::TypeEntry {
        name: name.into(),
        kind: vybe_runtime::chunk::CompositeKind::Struct,
        parent_index: 0,
        fields: Vec::new(),
        methods: Vec::new(),
        is_interface: false,
        implements: Vec::new(),
        constructor_chunk: None,
        field_descriptors: std::collections::HashMap::new() });
}

const TYPE_ONE: vybe_runtime::opcode::heaptype::HeapType =
    vybe_runtime::opcode::heaptype::HeapType::Concrete(1);

fn is_wasm_true(value: &Value) -> bool {
    matches!(value, Value::Bool(true) | Value::I32(1))
}

fn is_wasm_false(value: &Value) -> bool {
    matches!(value, Value::Bool(false) | Value::I32(0))
}

// ============================================================
// TypeRegistry unit tests
// ============================================================

#[test]
fn test_type_registration_and_lookup() {
    let mut vm = VM::new();

    // Object (type 0) is always pre-registered
    assert_eq!(vm.type_registry.get_id("Object"), Some(0));
    assert_eq!(vm.type_registry.get_id("object"), Some(0)); // case-insensitive

    // Register new types
    let list_id = vm
        .type_registry
        .register(TypeDef::new("List").with_parent(0));
    let dict_id = vm
        .type_registry
        .register(TypeDef::new("Dictionary").with_parent(0));

    assert!(list_id > 0);
    assert!(dict_id > 0);
    assert_ne!(list_id, dict_id);

    // Lookup by name
    assert_eq!(vm.type_registry.get_id("List"), Some(list_id));
    assert_eq!(vm.type_registry.get_id("list"), Some(list_id)); // case-insensitive
    assert_eq!(vm.type_registry.get_id("Dictionary"), Some(dict_id));

    // Unknown type
    assert_eq!(vm.type_registry.get_id("NonExistent"), None);
}

#[test]
fn test_subtype_checking_direct_parent() {
    let mut vm = VM::new();

    let control_id = vm
        .type_registry
        .register(TypeDef::new("Control").with_parent(0));
    let button_id = vm
        .type_registry
        .register(TypeDef::new("Button").with_parent(control_id));

    // Button is a subtype of Control
    assert!(vm.type_registry.is_subtype(button_id, control_id));
    // Button is a subtype of Object
    assert!(vm.type_registry.is_subtype(button_id, 0));
    // Control is a subtype of Object
    assert!(vm.type_registry.is_subtype(control_id, 0));
    // Everything is a subtype of itself
    assert!(vm.type_registry.is_subtype(button_id, button_id));
    assert!(vm.type_registry.is_subtype(control_id, control_id));
    // Object is NOT a subtype of Button
    assert!(!vm.type_registry.is_subtype(0, button_id));
    // Control is NOT a subtype of Button
    assert!(!vm.type_registry.is_subtype(control_id, button_id));
}

#[test]
fn test_subtype_checking_deep_chain() {
    let mut vm = VM::new();

    // Object -> Control -> TextBoxBase -> TextBox -> RichTextBox
    let control_id = vm
        .type_registry
        .register(TypeDef::new("Control").with_parent(0));
    let textbox_base_id = vm
        .type_registry
        .register(TypeDef::new("TextBoxBase").with_parent(control_id));
    let textbox_id = vm
        .type_registry
        .register(TypeDef::new("TextBox").with_parent(textbox_base_id));
    let rich_id = vm
        .type_registry
        .register(TypeDef::new("RichTextBox").with_parent(textbox_id));

    // RichTextBox is a subtype of everything above it
    assert!(vm.type_registry.is_subtype(rich_id, textbox_id));
    assert!(vm.type_registry.is_subtype(rich_id, textbox_base_id));
    assert!(vm.type_registry.is_subtype(rich_id, control_id));
    assert!(vm.type_registry.is_subtype(rich_id, 0)); // Object

    // Siblings are not subtypes of each other
    let label_id = vm
        .type_registry
        .register(TypeDef::new("Label").with_parent(control_id));
    assert!(!vm.type_registry.is_subtype(label_id, textbox_id));
    assert!(!vm.type_registry.is_subtype(textbox_id, label_id));
    // But both are subtypes of Control
    assert!(vm.type_registry.is_subtype(label_id, control_id));
    assert!(vm.type_registry.is_subtype(textbox_id, control_id));
}

#[test]
fn test_method_resolution_own_type() {
    let mut vm = VM::new();

    // Register a host function
    vm.register_host_fn("test", "listAdd", Box::new(|_, _| Value::Null));
    let host_idx = *vm
        .host_registry
        .get(&("test".to_string(), "listAdd".to_string()))
        .unwrap();

    // Register List type with an "add" method
    let list_id = vm.type_registry.register(
        TypeDef::new("List")
            .with_parent(0)
            .host_method("add", host_idx),
    );

    // Resolve method on List
    let method = vm.type_registry.resolve_method(list_id, "add");
    assert!(method.is_some());
    match method.unwrap() {
        Method::HostFn(idx) => assert_eq!(*idx, host_idx),
        _ => panic!("Expected HostFn"),
    }
}

#[test]
fn test_method_resolution_inherited() {
    let mut vm = VM::new();

    vm.register_host_fn(
        "test",
        "toString",
        Box::new(|_, _| Value::String(Arc::from("str"))),
    );
    let to_string_idx = *vm
        .host_registry
        .get(&("test".to_string(), "toString".to_string()))
        .unwrap();

    // Add toString to Object (type 0)
    vm.type_registry
        .add_host_method(0, "tostring", to_string_idx);

    // Register List inheriting from Object
    let list_id = vm
        .type_registry
        .register(TypeDef::new("List").with_parent(0));

    // List should inherit toString from Object
    let method = vm.type_registry.resolve_method(list_id, "tostring");
    assert!(method.is_some());
    match method.unwrap() {
        Method::HostFn(idx) => assert_eq!(*idx, to_string_idx),
        _ => panic!("Expected HostFn"),
    }
}

#[test]
fn test_method_resolution_override() {
    let mut vm = VM::new();

    vm.register_host_fn("test", "base_count", Box::new(|_, _| Value::F64(0.0)));
    let base_fn = *vm
        .host_registry
        .get(&("test".to_string(), "base_count".to_string()))
        .unwrap();
    vm.register_host_fn("test", "list_count", Box::new(|_, _| Value::F64(42.0)));
    let override_fn = *vm
        .host_registry
        .get(&("test".to_string(), "list_count".to_string()))
        .unwrap();

    // Object has a count method
    vm.type_registry.add_host_method(0, "count", base_fn);

    // List overrides count
    let list_id = vm.type_registry.register(
        TypeDef::new("List")
            .with_parent(0)
            .host_method("count", override_fn),
    );

    // List.count should resolve to the override
    let method = vm.type_registry.resolve_method(list_id, "count");
    match method.unwrap() {
        Method::HostFn(idx) => assert_eq!(*idx, override_fn),
        _ => panic!("Expected HostFn"),
    }

    // Object.count should still be the base
    let method = vm.type_registry.resolve_method(0, "count");
    match method.unwrap() {
        Method::HostFn(idx) => assert_eq!(*idx, base_fn),
        _ => panic!("Expected HostFn"),
    }
}

// ============================================================
// ref_test opcode tests
// ============================================================

fn make_typed_object(type_name: &str) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from(type_name)));
    Value::Object(Arc::new(std::sync::Mutex::new(obj)))
}

fn make_typed_object_with_id(type_id: usize, type_name: &str) -> Value {
    let mut obj = Object::new_typed(type_id);
    obj.properties
        .insert("__type".into(), Value::String(Arc::from(type_name)));
    Value::Object(Arc::new(std::sync::Mutex::new(obj)))
}

#[test]
fn test_ref_test_opcode_with_type_string() {
    // Create a VM with types registered
    let mut vm = VM::new();
    let control_id = vm
        .type_registry
        .register(TypeDef::new("Control").with_parent(0));
    let _button_id = vm
        .type_registry
        .register(TypeDef::new("Button").with_parent(control_id));

    // Build a chunk that: push a Button object, ref_test "control"
    let mut chunk = Chunk::new("<test>");
    // Push a Button object (using __type property)
    let obj = make_typed_object("Button");
    let const_idx = chunk.add_constant(obj);
    chunk.emit_op_u16(Op::CONST, const_idx, 0);
    // ref_test with "control" type name
    declare_type(&mut chunk, "control");
    chunk.emit_ref_type_op(Op::REF_TEST, TYPE_ONE, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(
        is_wasm_true(&result),
        "Button should be a subtype of Control, got {:?}",
        result
    );
}

#[test]
fn test_ref_test_opcode_with_type_id() {
    let mut vm = VM::new();
    let control_id = vm
        .type_registry
        .register(TypeDef::new("Control").with_parent(0));
    let button_id = vm
        .type_registry
        .register(TypeDef::new("Button").with_parent(control_id));

    // Build chunk with a typed object (type_id set)
    let mut chunk = Chunk::new("<test>");
    let obj = make_typed_object_with_id(button_id, "Button");
    let const_idx = chunk.add_constant(obj);
    chunk.emit_op_u16(Op::CONST, const_idx, 0);
    declare_type(&mut chunk, "control");
    chunk.emit_ref_type_op(Op::REF_TEST, TYPE_ONE, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(
        is_wasm_true(&result),
        "Button (type_id) should be a subtype of Control, got {:?}",
        result
    );
}

#[test]
fn test_ref_test_opcode_negative() {
    let mut vm = VM::new();
    let control_id = vm
        .type_registry
        .register(TypeDef::new("Control").with_parent(0));
    let _button_id = vm
        .type_registry
        .register(TypeDef::new("Button").with_parent(control_id));
    let _list_id = vm
        .type_registry
        .register(TypeDef::new("List").with_parent(0));

    // A List is NOT a Button
    let mut chunk = Chunk::new("<test>");
    let obj = make_typed_object("List");
    let const_idx = chunk.add_constant(obj);
    chunk.emit_op_u16(Op::CONST, const_idx, 0);
    declare_type(&mut chunk, "button");
    chunk.emit_ref_type_op(Op::REF_TEST, TYPE_ONE, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(
        is_wasm_false(&result),
        "List should NOT be a subtype of Button, got {:?}",
        result
    );
}

#[test]
fn test_ref_test_with_js_types_array() {
    // JS classes use __types array for inheritance chain
    let mut vm = VM::new();

    let mut chunk = Chunk::new("<test>");
    // Create an object with __types = ["Animal", "Dog"]
    let mut obj = Object::new();
    let types_arr = Object::new_array(vec![
        Value::String(Arc::from("Animal")),
        Value::String(Arc::from("Dog")),
    ]);
    obj.properties.insert(
        "__types".into(),
        Value::Object(Arc::new(std::sync::Mutex::new(types_arr))),
    );
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Dog")));

    let const_idx = chunk.add_constant(Value::Object(Arc::new(std::sync::Mutex::new(obj))));
    chunk.emit_op_u16(Op::CONST, const_idx, 0);
    declare_type(&mut chunk, "animal");
    chunk.emit_ref_type_op(Op::REF_TEST, TYPE_ONE, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(
        is_wasm_true(&result),
        "Dog (via __types) should match Animal, got {:?}",
        result
    );
}

#[test]
fn test_ref_test_primitives() {
    let mut vm = VM::new();

    // String is "string"
    let mut chunk = Chunk::new("<test>");
    let str_idx = chunk.add_constant(Value::String(Arc::from("hello")));
    chunk.emit_op_u16(Op::CONST, str_idx, 0);
    declare_type(&mut chunk, "string");
    chunk.emit_ref_type_op(Op::REF_TEST, TYPE_ONE, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(
        is_wasm_true(&result),
        "String should match 'string', got {:?}",
        result
    );
}

// ============================================================
// Component Model resource tests
// ============================================================

#[test]
fn test_resource_table_basic() {
    use vybe_runtime::ResourceTable;

    let mut table = ResourceTable::new();
    let handle = table.create(1, Value::String(Arc::from("data")));
    assert!(table.is_valid(handle));

    let borrowed = table.borrow(handle).unwrap();
    assert!(matches!(borrowed, Value::String(_)));

    table.release_borrow(handle);
    let dropped = table.drop_resource(handle).unwrap();
    assert!(matches!(dropped, Value::String(_)));
    assert!(!table.is_valid(handle));
}

#[test]
fn test_resource_table_borrow_prevents_drop() {
    use vybe_runtime::ResourceTable;

    let mut table = ResourceTable::new();
    let handle = table.create(1, Value::Null);
    let _ = table.borrow(handle);

    // Can't drop while borrowed
    assert!(table.drop_resource(handle).is_err());

    // Release, then drop succeeds
    table.release_borrow(handle);
    assert!(table.drop_resource(handle).is_ok());
}

#[test]
fn test_register_resource_in_type_registry() {
    let mut vm = VM::new();
    let control_id = vm
        .type_registry
        .register(TypeDef::new("Control").with_parent(0));

    // Register a resource type with parent
    let mut td = TypeDef::new("FileHandle");
    td.parent = Some(control_id);
    td.is_resource = true;
    let file_tid = vm.type_registry.register(td);

    assert!(file_tid > 0);
    assert_eq!(vm.type_registry.get_id("FileHandle"), Some(file_tid));
    assert!(vm.type_registry.is_subtype(file_tid, control_id));
    assert!(vm.type_registry.is_subtype(file_tid, 0)); // Object
}
