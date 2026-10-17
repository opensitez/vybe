use std::path::Path;

#[test]
fn global_record_field_owner_matches_native_and_static_wasm() {
    assert_record_fixture("global_record_field_owner.c");
}

#[test]
fn record_pointer_field_owner_matches_native_and_static_wasm() {
    assert_record_fixture("record_pointer_field_owner.c");
}

#[test]
fn config_collection_lookup_preserves_backed_record_names_and_bindings() {
    assert_record_fixture("config_defaults_collection_search.c");
}

#[test]
fn player_reborn_preserves_inline_frags_across_save_zero_restore() {
    assert_record_fixture("doom_player_reborn_frags.c");
}

#[test]
fn nested_record_array_zero_preserves_backing_aliases_and_pointer_slots() {
    assert_record_fixture("nested_record_array_zero_all.c");
}

fn assert_record_fixture(name: &str) {
    vybe_language_c::register();
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/c/doom_startup");
    let source = root.join(name);
    let native = std::env::temp_dir().join(format!(
        "vybe-global-owner-{}-{}",
        std::process::id(),
        name.trim_end_matches(".c")
    ));
    let compiled = std::process::Command::new("cc")
        .args(["-std=c11", "-O0"])
        .arg(&source)
        .arg("-o")
        .arg(&native)
        .output()
        .expect("native C compiler");
    assert!(
        compiled.status.success(),
        "{name}: {}",
        String::from_utf8_lossy(&compiled.stderr)
    );
    let executed = std::process::Command::new(&native).output().unwrap();
    assert!(
        executed.status.success(),
        "{name}: {}",
        String::from_utf8_lossy(&executed.stderr)
    );
    std::fs::remove_file(&native).unwrap();

    let chunks = vybe_compiler::projects::load(&source)
        .unwrap()
        .compile()
        .unwrap();
    for chunk in &chunks {
        assert!(!chunk.name.contains("dyneq"), "{name}: {}", chunk.name);
        assert!(
            !vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "{name}: {}",
            chunk.name
        );
    }
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: false,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    let mut abstract_imports = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(section) => {
                for import in section.into_imports() {
                    let import = import.unwrap();
                    if matches!(
                        import.ty,
                        wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)
                    ) {
                        abstract_imports.push(
                            import.module == "ecma:value"
                                && matches!(import.name, "abstractEq" | "abstractNe"),
                        );
                    }
                }
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                let mut operators = body.get_operators_reader().unwrap();
                while !operators.eof() {
                    if let wasmparser::Operator::Call { function_index }
                    | wasmparser::Operator::ReturnCall { function_index }
                    | wasmparser::Operator::RefFunc { function_index } =
                        operators.read().unwrap()
                    {
                        assert_ne!(
                            abstract_imports.get(function_index as usize),
                            Some(&true),
                            "{name}"
                        );
                    }
                }
            }
            _ => {}
        }
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.memory.grow(1);
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0, "{name}");
}
