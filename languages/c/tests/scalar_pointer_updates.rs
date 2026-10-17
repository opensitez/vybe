#[test]
fn named_integer_pointer_updates_preserve_values_and_validate_as_wasm() {
    vybe_language_c::register();
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup");
    for name in ["scalar_pointer_integer_updates.c", "hud_ammo_pointer_alias.c",
        "doom_pistol_ammo_update.c"] {
        let chunks = vybe_compiler::projects::load(&root.join(name)).unwrap().compile().unwrap();
        for chunk in &chunks {
            assert!(!vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
                "{name}/{}: dynamic equality", chunk.name);
        }
        let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false, include_name_section: true,
            });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap();
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm,
            &vybe_runtime::capabilities::Capabilities::all());
        vm.run(chunks).unwrap();
        assert_eq!(vm.pending_exit_code, 0, "{name}");
    }
}
