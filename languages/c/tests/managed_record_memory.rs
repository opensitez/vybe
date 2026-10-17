#[test]
fn complete_record_zero_fill_preserves_aliases_and_matches_native() {
    check_fixture("managed_record_memset_reset.c");
}

#[test]
fn local_tagged_record_byte_arrays_preserve_layout_and_unsigned_values() {
    check_fixture("local_tagged_record_bytes.c");
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
    eprintln!("{name}: noEq PASS {} chunks; metadata-free WASM valid {} bytes", chunks.len(), bytes.len());
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm,
        &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
