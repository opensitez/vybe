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
    vybe_platform_wasm::validate_binary(&bytes).unwrap();
    let mut abstract_imports = Vec::new();
    let module = vybe_platform_wasm::inspection::CoreModule::new(&bytes).unwrap();
    module.imports(|module, field, ty| {
        if matches!(ty, vybe_platform_wasm::inspection::ImportType::Func { .. }) {
            abstract_imports.push(module == "ecma:value"
                && matches!(field, "abstractEq" | "abstractNe"));
        }
        Ok(())
    }).unwrap();
    module.function_bodies(|body| {
        for instruction in body.instructions()? {
            if let Some(index) = instruction?.function_index() {
                assert_ne!(abstract_imports.get(index as usize), Some(&true), "{name}");
            }
        }
        Ok(())
    }).unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.memory.grow(1);
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0, "{name}");
}
