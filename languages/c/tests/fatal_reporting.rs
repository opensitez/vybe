use std::sync::{Arc, Mutex};

fn compile_fixture(name: &str) -> Vec<vybe_runtime::Chunk> {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup")
        .join(name);
    let chunks = vybe_compiler::projects::load(&path)
        .unwrap()
        .compile()
        .unwrap();
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    chunks
}

#[test]
fn exit_terminates_through_nested_callers() {
    let chunks = compile_fixture("fatal_exit_nested.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 7);
}

#[test]
fn reentrant_fatal_handler_cannot_resume_its_outer_call() {
    let chunks = compile_fixture("fatal_exit_reentrant.c");
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 9);
}

#[test]
fn sdl_error_dialog_decodes_c_buffers_and_literals() {
    let chunks = compile_fixture("sdl_error_message.c");
    let messages = Arc::new(Mutex::new(Vec::new()));
    let captured = messages.clone();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.register_host_fn(
        "web:window",
        "alert",
        Box::new(move |_, args| {
            let vybe_runtime::Value::String(message) = &args[0] else {
                panic!("alert must receive text, got {:?}", args[0]);
            };
            captured.lock().unwrap().push(message.to_string());
            vybe_runtime::Value::Null
        }),
    );
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    assert_eq!(
        *messages.lock().unwrap(),
        vec![
            "Chocolate Doom 3.1.1\n\nBad patch: 42",
            "Literal title\n\nLinear message",
            "Doom 3.1.1\n\nmessage",
        ]
    );
}
