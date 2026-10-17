use vybe_platform_wasm::read_wasm;
use vybe_runtime::VM;

#[test]
fn native_descriptor_allocation_cast_and_vtable_module_loads_and_runs() {
    let bytes = include_bytes!("../../../crates/vybe_runtime/tests/fixtures/custom_descriptor_classes.wasm");
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(bytes)
        .expect("fixture must validate against standard proposal operators");
    let chunks = read_wasm(bytes).expect("Vybe must accept the valid descriptor stack effects");
    // The module's start function traps on a failed type test, identity check,
    // descriptor cast, or descriptor-vtable call.
    VM::new().run(chunks).expect("descriptor class checks must execute successfully");
}
