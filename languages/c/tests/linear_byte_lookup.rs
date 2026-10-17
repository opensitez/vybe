use std::path::Path;
use std::sync::atomic::Ordering;
use vybe_runtime::{DebugCommand, DebugRequest, VM};

#[test]
fn generated_linear_byte_lookup_preserves_values_without_repeated_narrowing() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/linear_byte_lookup.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let chunks = bundle.compile().unwrap();
    let lookup = chunks.iter().find(|chunk| chunk.name == "lookup").unwrap();
    let dump = vybe_runtime::debug::disassemble(lookup);
    assert_eq!(dump.matches("i32.load8_u").count(), 2, "{dump}");
    assert!(!dump.contains("i32.and"), "{dump}");

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

    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let (commands, rx) = std::sync::mpsc::channel();
    let (events, _rx) = std::sync::mpsc::channel();
    vm.attach_debugger(rx, events, false);
    let (reply, _response) = std::sync::mpsc::channel();
    commands
        .send(DebugRequest {
            command: DebugCommand::Continue,
            reply,
        })
        .unwrap();
    let report = vm.debug_report().unwrap();
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    eprintln!(
        "linear byte lookup VM instructions: {}",
        report
            .lock()
            .unwrap()
            .live
            .instructions
            .load(Ordering::Relaxed)
    );
}
