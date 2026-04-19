/// Tests for Component Model: canon_lift, canon_lower, type_import, type_export.

use vybe_bytecode::{VM, Value, Chunk, Op, TypeDef};
use std::rc::Rc;

#[test]
fn canon_lift_stamps_type_id() {
    let mut vm = VM::new();
    let mut td = TypeDef::new("Animal");
    td.add_field("name");
    let tid = vm.type_registry.register(td);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Create a plain object (type_id = 0)
    let name_c = chunk.add_constant(Value::String(Rc::from("name")));
    let val_c = chunk.add_constant(Value::String(Rc::from("Rex")));
    chunk.emit_op_u16(Op::CONST, val_c, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 1, 0);

    // canon_lift with Animal type
    chunk.emit_op_u16(Op::CANON_LIFT, tid as u16, 0);

    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    match &result {
        Value::Object(obj) => {
            let o = obj.borrow();
            assert_eq!(o.type_id, tid, "canon_lift should stamp type_id");
        }
        other => panic!("expected Object, got {:?}", other),
    }
}

#[test]
fn canon_lower_passes_through() {
    let mut vm = VM::new();
    let tid = vm.type_registry.register(TypeDef::new("MyType"));

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    // Create typed object
    let type_id = chunk.add_constant(Value::I32(tid as i32));
    chunk.emit_op_u16(Op::CONST, type_id, 0);
    chunk.emit_op(Op::SHARED_NEW, 0);

    // canon_lower
    chunk.emit_op_u16(Op::CANON_LOWER, tid as u16, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn type_import_resolves_registered_type() {
    let mut vm = VM::new();
    let tid = vm.type_registry.register(TypeDef::new("Widget"));

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.type_imports.push(("ui".to_string(), "Widget".to_string()));

    // type_import 0 → should push Widget's type_id
    chunk.emit_op_u16(Op::TYPE_IMPORT, 0, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), tid as i32);
}

#[test]
fn type_import_unknown_returns_null() {
    let mut vm = VM::new();

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.type_imports.push(("pkg".to_string(), "NonExistent".to_string()));

    chunk.emit_op_u16(Op::TYPE_IMPORT, 0, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Null));
}

#[test]
fn type_export_is_noop() {
    let mut vm = VM::new();
    let tid = vm.type_registry.register(TypeDef::new("Exported"));

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    // type_export just reads operand, no stack effect
    chunk.emit_op_u16(Op::TYPE_EXPORT, tid as u16, 0);
    let val = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 1);
}

// ── Linker type resolution ──────────────────────────────────

#[test]
fn linker_resolves_type_exports() {
    use vybe_bytecode::component::*;
    use std::collections::HashMap;

    let mut td = TypeDef::new("Dog");
    td.add_field("name");
    td.interface = Some("animals:api".to_string());

    let mut comp_a = Component {
        name: "animal-lib".into(),
        language: Language::Python,
        chunks: vec![Chunk::new("<script>")],
        imports: vec![],
        exports: HashMap::new(),
        type_exports: {
            let mut m = HashMap::new();
            m.insert(("animals:api".into(), "Dog".into()), td);
            m
        },
        type_imports: vec![],
    };

    let comp_b = Component {
        name: "app".into(),
        language: Language::JS,
        chunks: vec![Chunk::new("<script>")],
        imports: vec![],
        exports: HashMap::new(),
        type_exports: HashMap::new(),
        type_imports: vec![("animals:api".into(), "Dog".into())],
    };

    let mut linker = Linker::new();
    linker.add_component(comp_a);
    linker.add_component(comp_b);

    let result = linker.link().expect("linking should succeed");
    // Type export should be in the result
    assert!(result.type_exports.contains_key(&("animals:api".into(), "Dog".into())));
}

// ── TypeRegistry import/export ──────────────────────────────

#[test]
fn type_registry_import_type() {
    let mut source_td = TypeDef::new("Cat");
    source_td.add_field("name");
    source_td.add_field("age");
    source_td.interface = Some("pets:api".to_string());

    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let id = reg.import_type(&source_td);

    assert!(id > 0);
    let imported = reg.get(id).unwrap();
    assert_eq!(imported.name, "Cat");
    assert_eq!(imported.field_defs.len(), 2);
    assert_eq!(imported.interface.as_deref(), Some("pets:api"));
}

#[test]
fn type_registry_import_merges_existing() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();

    // First registration
    let mut td1 = TypeDef::new("Dog");
    td1.add_field("name");
    let id1 = reg.register(td1);

    // Import same type with additional field
    let mut td2 = TypeDef::new("Dog");
    td2.add_field("name");
    td2.add_field("breed");
    let id2 = reg.import_type(&td2);

    assert_eq!(id1, id2, "should reuse existing type");
    let merged = reg.get(id1).unwrap();
    assert_eq!(merged.field_defs.len(), 2, "should have merged fields");
}

#[test]
fn type_registry_resolve_type_import() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let mut td = TypeDef::new("Button");
    td.interface = Some("gui:controls".to_string());
    reg.register(td);

    let resolved = reg.resolve_type_import("gui:controls", "Button");
    assert!(resolved.is_some());
}

#[test]
fn type_registry_export_and_query() {
    let mut reg = vybe_bytecode::typedef::TypeRegistry::new();
    let tid = reg.register(TypeDef::new("Form"));
    reg.export_type(tid, "gui:forms", "my-app");

    let exports = reg.get_component_exports("my-app");
    assert_eq!(exports.len(), 1);
    assert_eq!(exports[0].1.name, "Form");
}
