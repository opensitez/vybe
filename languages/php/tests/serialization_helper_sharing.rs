use std::sync::Arc;
use vybe_compiler::primitives::helper_templates::{Context, Owner, Pool};
use vybe_runtime::{Chunk, Value, chunk::TypeEntry};

fn compiled(source: &str, selected: bool) -> Vec<Chunk> {
    vybe_language_php::register();
    let ast = vybe_language_php::parse(source).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    let compiler = vybe_compiler::Compiler::with_profile(profile);
    let compiler = if selected {
        compiler.with_helper_templates(
            vybe_language_php::parse,
            vybe_language_php::emitter::serialization_adapter::HELPER_DEFINITIONS,
            Arc::new(Pool::default()),
        )
    } else {
        compiler
    };
    compiler.compile(&ast).unwrap()
}

fn helper_count(chunks: &[Chunk]) -> usize {
    chunks
        .iter()
        .filter(|chunk| chunk.name == "__php_serialize_value")
        .count()
}

#[test]
fn included_units_preserve_namespaces_aliases_and_identical_short_names() {
    vybe_language_php::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/shared_helper_namespaces/main.php");
    let result = vybe_compiler::dynamic::RuntimeCompilerService::with_capabilities(
        &mut vm,
        vybe_runtime::capabilities::Capabilities::all(),
    )
    .compile_and_run_path(&path)
    .unwrap();
    assert_eq!(
        result,
        Value::String("first-label:first|second-label:second|HelperConsumer".into())
    );
    assert!(vm.last_exception.is_none());
}

#[test]
fn repeated_serialization_builds_one_owned_recursive_body_and_preserves_results() {
    let source = "<?php $value = ['x'=>[1,2]]; $a=serialize($value); $b=serialize($value); return $a.'|'.$b;";
    let before = compiled(source, false);
    let after = compiled(source, true);
    assert_eq!(helper_count(&before), 2);
    assert_eq!(helper_count(&after), 1);
    let bytes = |chunks: &[Chunk]| chunks.iter().map(|chunk| chunk.code.len()).sum::<usize>();
    assert!(bytes(&after) < bytes(&before));
    let owned_index = after
        .iter()
        .position(|chunk| chunk.name == "__php_serialize_value")
        .unwrap();
    assert!(!after[owned_index].direct_defined_calls.is_empty());
    assert!(
        after[owned_index]
            .direct_defined_calls
            .values()
            .all(|edge| edge.target as usize == owned_index)
    );
    let mut values = Vec::new();
    for chunks in [before, after] {
        let wasm = vybe_platform_wasm::write_wasm_with_options(
            &chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&wasm)
            .unwrap();
        let mut vm = vybe_runtime::VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        values.push(vm.run(chunks).unwrap());
    }
    assert_eq!(values[0], values[1]);
    assert!(matches!(&values[0], Value::String(value) if value.contains("a:1:")));
}

#[test]
fn separate_source_lines_keep_separate_helper_diagnostics() {
    let chunks = compiled(
        "<?php $x=[1];\n$a=serialize($x);\nreturn serialize($x);",
        true,
    );
    assert_eq!(helper_count(&chunks), 2);
}

#[test]
fn repeated_unserialization_shares_allocator_and_decoder_and_preserves_results() {
    for options in ["", ", []"] {
        let source = format!(
            "<?php $wire='a:2:{{i:0;i:3;i:1;s:1:\"x\";}}'; $a=unserialize($wire{options}); $b=unserialize($wire{options}); return serialize([$a,$b]);"
        );
        let before = compiled(&source, false);
        let after = compiled(&source, true);
        for name in ["__php_unserialize_alloc", "__php_unserialize_value"] {
            assert_eq!(before.iter().filter(|chunk| chunk.name == name).count(), 2);
            assert_eq!(after.iter().filter(|chunk| chunk.name == name).count(), 1);
        }
        let bytes = |chunks: &[Chunk]| chunks.iter().map(|chunk| chunk.code.len()).sum::<usize>();
        assert!(bytes(&after) < bytes(&before));
        let mut values = Vec::new();
        for chunks in [before, after] {
            let wasm = vybe_platform_wasm::write_wasm_with_options(
                &chunks,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false,
                    include_name_section: true,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&wasm)
                .unwrap();
            let mut vm = vybe_runtime::VM::new();
            vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
            values.push(vm.run(chunks).unwrap());
        }
        assert_eq!(values[0], values[1]);
        assert!(matches!(&values[0], Value::String(value)
            if value.as_ref() == "a:2:{i:0;a:2:{i:0;i:3;i:1;s:1:\"x\";}i:1;a:2:{i:0;i:3;i:1;s:1:\"x\";}}"));
    }
}

#[test]
fn unserialize_provider_tracks_type_positions_fields_methods_and_source_lines() {
    let profile =
        vybe_runtime::profile::parse_profile_shared(vybe_language_php::profile_source()).unwrap();
    let directives = vybe_ast::Directives::default();
    let context = Context {
        profile: &profile,
        directives: &directives,
        receiver: Default::default(),
        own: false,
        prototype: false,
    };
    let mut owner = Owner::new(
        vybe_language_php::emitter::serialization_adapter::HELPER_DEFINITIONS,
        Arc::new(Pool::default()),
    );
    let mut chunks = vec![Chunk::new("caller"), Chunk::new("__php_unserialize_value")];
    let emit = |owner: &mut Owner, chunks: &mut Vec<Chunk>, line| {
        assert!(owner.emit(
            "php.unserialize",
            vybe_language_php::parse,
            vybe_language_php::emitter::dispatch::dispatch,
            1,
            vybe_runtime::registry::emitter_results::EmitRequest::Value,
            &context,
            chunks,
            0,
            line
        ));
    };
    emit(&mut owner, &mut chunks, 17);
    assert_eq!(chunks.len(), 4, "guest decoder is never adopted");
    emit(&mut owner, &mut chunks, 17);
    assert_eq!(chunks.len(), 4);
    chunks[0].types.push(TypeEntry {
        name: "Added".into(),
        fields: vec!["field".into()],
        ..Default::default()
    });
    emit(&mut owner, &mut chunks, 17);
    assert_eq!(chunks.len(), 6);
    chunks[0].types[0].fields.push("__type".into());
    emit(&mut owner, &mut chunks, 17);
    assert_eq!(
        chunks.len(),
        8,
        "allocator consumes all fields, including internal fields"
    );
    chunks[0].types[0].methods.push(("method".into(), 1));
    emit(&mut owner, &mut chunks, 17);
    assert_eq!(chunks.len(), 10);
    chunks[0].types.insert(
        0,
        TypeEntry {
            name: "Interface".into(),
            is_interface: true,
            ..Default::default()
        },
    );
    emit(&mut owner, &mut chunks, 17);
    assert_eq!(
        chunks.len(),
        12,
        "interface insertion shifts concrete type indices"
    );
    emit(&mut owner, &mut chunks, 18);
    assert_eq!(
        chunks.len(),
        14,
        "source diagnostics retain their original line"
    );
}

#[test]
fn catalog_and_owner_changes_rebuild_and_guest_name_is_never_adopted() {
    let profile =
        vybe_runtime::profile::parse_profile_shared(vybe_language_php::profile_source()).unwrap();
    let directives = vybe_ast::Directives::default();
    let context = Context {
        profile: &profile,
        directives: &directives,
        receiver: Default::default(),
        own: false,
        prototype: false,
    };
    let definitions = vybe_language_php::emitter::serialization_adapter::HELPER_DEFINITIONS;
    let mut owner = Owner::new(definitions, Arc::new(Pool::default()));
    let mut chunks = vec![Chunk::new("caller"), Chunk::new("__php_serialize_value")];
    let emit = |owner: &mut Owner, chunks: &mut Vec<Chunk>| {
        owner.emit(
            "php.serialize",
            vybe_language_php::parse,
            vybe_language_php::emitter::dispatch::dispatch,
            1,
            vybe_runtime::registry::emitter_results::EmitRequest::Value,
            &context,
            chunks,
            0,
            17,
        )
    };
    assert!(emit(&mut owner, &mut chunks));
    assert_eq!(chunks.len(), 3);
    assert!(emit(&mut owner, &mut chunks));
    assert_eq!(chunks.len(), 3);
    chunks[0].types.push(TypeEntry {
        name: "Added".into(),
        fields: vec!["field".into()],
        ..Default::default()
    });
    assert!(emit(&mut owner, &mut chunks));
    assert_eq!(
        chunks.len(),
        4,
        "changed class catalog cannot reuse an older serializer"
    );
    chunks[0].types[0].fields.push("later".into());
    assert!(emit(&mut owner, &mut chunks));
    assert_eq!(
        chunks.len(),
        5,
        "changed ordered fields cannot reuse old code"
    );
    let mut other = Owner::new(definitions, Arc::new(Pool::default()));
    assert!(emit(&mut other, &mut chunks));
    assert_eq!(
        chunks.len(),
        6,
        "another owner cannot adopt prior or guest helpers by name"
    );
}
