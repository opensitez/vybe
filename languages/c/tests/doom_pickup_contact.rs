#[test]
fn normal_player_pickup_contact_preserves_health_caps_and_record_backing() {
    vybe_language_c::register();
    let source = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/doom_pickup_contact.c");
    let chunks = vybe_compiler::projects::load(&source).unwrap().compile().unwrap();
    for chunk in &chunks {
        assert!(!vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "dynamic equality in {}", chunk.name);
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
    assert_eq!(vm.pending_exit_code, 0);
}
