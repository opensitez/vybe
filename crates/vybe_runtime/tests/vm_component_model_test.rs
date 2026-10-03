/// Tests for Component Model: canon lift, canon lower, type_import, type_export.
use vybe_runtime::{Chunk, TypeDef, VM, Value};

#[test]
fn bare_canon_lift_requires_canon_section_row() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");

    // `canon lift` is not a type-stamping stub. It is an instantiation-time
    // built-in produced by a `(canon lift ...)` row, so the import must name
    // that row as `lift@<canonidx>`.
    let lift = chunk.add_import("canon", "lift");
    chunk.emit_call(lift, 0, 0);

    let err = vm.run(vec![chunk]).unwrap_err();
    assert!(err.message.contains("canon lift: no canonidx"));
    assert!(err.message.contains("lift@<canonidx>"));
}

// ── Linker type resolution ──────────────────────────────────

#[test]
fn linker_resolves_type_exports() {
    use std::collections::HashMap;
    use vybe_runtime::component::*;

    let mut td = TypeDef::new("Dog");
    td.add_field("name");
    td.interface = Some("animals:api".to_string());

    let comp_a = Component {
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
    assert!(
        result
            .type_exports
            .contains_key(&("animals:api".into(), "Dog".into()))
    );
}

#[test]
fn linker_register_host_from_vm_uses_module_records() {
    use std::collections::HashMap;
    use vybe_runtime::ImportTarget;
    use vybe_runtime::chunk::Import;
    use vybe_runtime::component::{Component, Language, Linker};

    let mut vm = VM::new();
    vm.register_host_fn("ecma:test", "alpha", Box::new(|_ctx, _args| Value::I32(1)));
    vm.register_host_fn("wasi:test", "beta", Box::new(|_ctx, _args| Value::I32(2)));

    vm.host_registry
        .remove(&("ecma:test".to_string(), "alpha".to_string()));
    vm.host_registry
        .remove(&("wasi:test".to_string(), "beta".to_string()));

    let mut script = Chunk::new("<script>");
    script.imports.push(Import {
        module: "ecma:test".into(),
        name: "alpha".into(),
    });
    script.imports.push(Import {
        module: "wasi:test".into(),
        name: "beta".into(),
    });

    let component = Component {
        name: "app".into(),
        language: Language::JS,
        chunks: vec![script],
        imports: vec![
            ("ecma:test".into(), "alpha".into()),
            ("wasi:test".into(), "beta".into()),
        ],
        exports: HashMap::new(),
        type_exports: HashMap::new(),
        type_imports: vec![],
    };

    let mut linker = Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(component);

    let result = linker
        .link()
        .expect("host imports should resolve from module records");
    assert!(matches!(
        result.resolved_imports.first(),
        Some(ImportTarget::Host(_))
    ));
    assert!(matches!(
        result.resolved_imports.get(1),
        Some(ImportTarget::Host(_))
    ));
}

#[test]
fn linker_register_host_from_vm_includes_host_type_exports() {
    use std::collections::HashMap;
    use vybe_runtime::component::{Component, Language, Linker};

    let mut vm = VM::new();
    let mut descriptor = TypeDef::new("Descriptor");
    descriptor.interface = Some("wasi:filesystem/types".into());
    descriptor.is_resource = true;
    let tid = vm.type_registry.register(descriptor);
    vm.register_host_resource_type_export("wasi:filesystem/types", "descriptor", tid);

    let component = Component {
        name: "fs-client".into(),
        language: Language::JS,
        chunks: vec![Chunk::new("<script>")],
        imports: vec![],
        exports: HashMap::new(),
        type_exports: HashMap::new(),
        type_imports: vec![("wasi:filesystem/types".into(), "descriptor".into())],
    };

    let mut linker = Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(component);

    let result = linker
        .link()
        .expect("host type imports should resolve from module records");
    let exported = result
        .type_exports
        .get(&("wasi:filesystem/types".into(), "descriptor".into()))
        .expect("host type export should be visible to the linker");
    assert_eq!(exported.name, "Descriptor");
    assert!(exported.is_resource);
}

#[test]
fn esm_and_component_linker_share_canonical_host_subinterface_aliases() {
    use std::collections::HashMap;
    use vybe_runtime::ImportTarget;
    use vybe_runtime::component::{Component, Language, Linker};

    let mut vm = VM::new();
    vm.register_host_fn(
        "node:util",
        "types.isArray",
        Box::new(|_ctx, _args| Value::I32(7)),
    );
    vm.host_registry
        .remove(&("node:util".to_string(), "types.isArray".to_string()));

    let mut script = Chunk::new("<script>");
    script.local_count = 0;
    let import_idx = script.add_import("node:util/types", "isArray");
    script.emit_call(import_idx, 0, 0);

    let vm_result = vm
        .run(vec![script.clone()])
        .expect("VM should resolve the canonical host subinterface alias");
    assert_eq!(vm_result.as_i32(), 7);

    let component = Component {
        name: "app".into(),
        language: Language::JS,
        chunks: vec![script],
        imports: vec![("node:util/types".into(), "isArray".into())],
        exports: HashMap::new(),
        type_exports: HashMap::new(),
        type_imports: vec![],
    };

    let mut linker = Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(component);

    let link_result = linker
        .link()
        .expect("Component linker should resolve the same canonical host subinterface alias");
    assert!(matches!(
        link_result.resolved_imports.first(),
        Some(ImportTarget::Host(_))
    ));

    let mut linked_vm = VM::new();
    linked_vm.register_host_fn(
        "node:util",
        "types.isArray",
        Box::new(|_ctx, _args| Value::I32(7)),
    );
    linked_vm
        .host_registry
        .remove(&("node:util".to_string(), "types.isArray".to_string()));

    let linked_result = linked_vm
        .run_linked(link_result.chunks, link_result.resolved_imports)
        .expect("linked execution should use the same host function target");
    assert_eq!(linked_result.as_i32(), 7);
}

// ── TypeRegistry import/export ──────────────────────────────

#[test]
fn type_registry_import_type() {
    let mut source_td = TypeDef::new("Cat");
    source_td.add_field("name");
    source_td.add_field("age");
    source_td.interface = Some("pets:api".to_string());

    let mut reg = vybe_runtime::typedef::TypeRegistry::new();
    let id = reg.import_type(&source_td);

    assert!(id > 0);
    let imported = reg.get(id).unwrap();
    assert_eq!(imported.name, "Cat");
    assert_eq!(imported.field_defs.len(), 2);
    assert_eq!(imported.interface.as_deref(), Some("pets:api"));
}

#[test]
fn type_registry_import_merges_existing() {
    let mut reg = vybe_runtime::typedef::TypeRegistry::new();

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
fn type_registry_export_and_query() {
    let mut reg = vybe_runtime::typedef::TypeRegistry::new();
    let tid = reg.register(TypeDef::new("Form"));
    reg.export_type(tid, "gui:forms", "my-app");

    let exports = reg.get_component_exports("my-app");
    assert_eq!(exports.len(), 1);
    assert_eq!(exports[0].1.name, "Form");
}
