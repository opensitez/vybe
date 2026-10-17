use std::collections::HashMap;
use std::path::Path;

#[test]
fn byte_lookup_views_retain_live_storage_and_descriptor_updates() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/byte_lookup_view_stability.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let module = bundle.prepared_module().unwrap();
    let chunks = bundle.compile_prepared_module(&module, &HashMap::new()).unwrap().chunks;
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
