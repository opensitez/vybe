#[test]
fn scalar_cell_addresses_round_trip_through_direct_and_nested_record_fields() {
    check_fixture("scalar_cell_record_pointer.c");
}

#[test]
fn scalar_cell_addresses_survive_record_copy_move_and_zero() {
    check_fixture("scalar_cell_record_pointer_memory.c");
}

#[test]
fn large_record_copy_preserves_scalar_addresses_and_nested_aliases() {
    check_fixture("managed_large_record_copy.c");
}

#[test]
fn nonidentifier_integer_pointer_updates_evaluate_the_pointer_once() {
    check_fixture("nonident_integer_pointer_updates.c");
}

fn check_fixture(name: &str) {
    vybe_language_c::register();
    let source = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup").join(name);
    let native = std::env::temp_dir().join(format!("vybe-{name}-{}", std::process::id()));
    assert!(std::process::Command::new("cc").arg(&source).arg("-o").arg(&native)
        .status().unwrap().success());
    assert!(std::process::Command::new(&native).status().unwrap().success());
    let _ = std::fs::remove_file(&native);
    let chunks = vybe_compiler::projects::load(&source).unwrap().compile().unwrap();
    for chunk in &chunks {
        assert!(!vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "dynamic equality in {}", chunk.name);
        let imports = if chunk.imports.is_empty() { &chunks[0].imports } else { &chunk.imports };
        let mut offset = 0;
        while offset + 4 <= chunk.code.len() {
            let op = vybe_runtime::opcode::Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2)).unwrap();
            if op == vybe_runtime::opcode::Op::CALL {
                if let Some(import) = imports.get(chunk.read_u16(offset + 4) as usize) {
                    assert!(import.module != "ecma:value"
                        || !matches!(import.name.as_str(), "abstractEq" | "abstractNe"));
                }
            }
            let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
            assert!(next > offset);
            offset = next;
        }
    }
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    eprintln!("scalar-cell records: noEq PASS {} chunks; metadata-free WASM valid {} bytes",
        chunks.len(), bytes.len());
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm,
        &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
