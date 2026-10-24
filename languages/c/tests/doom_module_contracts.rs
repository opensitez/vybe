use std::path::Path;
use vybe_runtime::opcode::Op;

#[test]
#[ignore = "compiles the complete Doom translation unit"]
fn complete_doom_has_no_dynamic_equality_and_validates_as_wasm() {
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let source = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../examples/chocolate-doom/vybe_doom_core.c");
    let started = std::time::Instant::now();
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_path(&source)
        .unwrap();
    eprintln!(
        "Doom compilation: {:?}, {} chunks",
        started.elapsed(),
        compiled.chunks.len()
    );
    drop(vm);
    let audit_started = std::time::Instant::now();
    let mut violations = Vec::new();
    for chunk in &compiled.chunks {
        if chunk.name.contains("dyneq") {
            violations.push(format!("{}: dynamic equality helper", chunk.name));
        }
        let imports = if chunk.imports.is_empty() {
            &compiled.chunks[0].imports
        } else {
            &chunk.imports
        };
        let mut offset = 0;
        while offset < chunk.code.len() {
            assert!(
                offset + 4 <= chunk.code.len(),
                "{}@{offset}: truncated opcode",
                chunk.name
            );
            let op = Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2)).unwrap();
            let (text, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
            if text.contains("dyneq") {
                violations.push(format!("{}@{offset}: {text}", chunk.name));
            }
            assert!(
                next > offset && next <= chunk.code.len(),
                "{}@{offset}: invalid instruction boundary {next}",
                chunk.name
            );
            if op == Op::CALL {
                let index = chunk.read_u16(offset + 4) as usize;
                if let Some(import) = imports.get(index) {
                    if import.module == "ecma:value"
                        && matches!(import.name.as_str(), "abstractEq" | "abstractNe")
                    {
                        violations.push(format!(
                            "{}@{offset}: {}:{}",
                            chunk.name, import.module, import.name
                        ));
                    }
                }
            }
            offset = next;
        }
    }
    assert!(
        violations.is_empty(),
        "dynamic equality in full Doom:\n{}",
        violations.join("\n")
    );
    eprintln!(
        "Doom dynamic-equality audit PASS: {:?}",
        audit_started.elapsed()
    );
    let write_started = std::time::Instant::now();
    let wasm = vybe_platform_wasm::write_wasm_with_options(
        &compiled.chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    eprintln!(
        "Doom WASM serialization: {:?}, {} bytes",
        write_started.elapsed(),
        wasm.len()
    );
    drop(compiled);
    let validation_started = std::time::Instant::now();
    vybe_platform_wasm::validate_binary(&wasm)
        .unwrap();
    audit_serialized_equality(&wasm);
    eprintln!(
        "Doom metadata-free WASM validation PASS: {:?}",
        validation_started.elapsed()
    );
}

fn audit_serialized_equality(wasm: &[u8]) {
    let module = vybe_platform_wasm::inspection::CoreModule::new(wasm).unwrap();
    let mut abstract_imports = Vec::new();
    module.imports(|module, name, ty| {
        if matches!(ty, vybe_platform_wasm::inspection::ImportType::Func { .. }) {
            abstract_imports.push(module == "ecma:value"
                && matches!(name, "abstractEq" | "abstractNe"));
        }
        Ok(())
    }).unwrap();
    module.function_bodies(|body| {
        for instruction in body.instructions()? {
            let instruction = instruction?;
            if let Some(index) = instruction.function_index() {
                assert_ne!(abstract_imports.get(index as usize), Some(&true),
                    "abstract equality call/reference in serialized WASM at {}", instruction.offset);
            }
        }
        Ok(())
    }).unwrap();
}
