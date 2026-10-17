use std::path::Path;

#[test]
fn typed16_for_column_matches_scalar_native_and_validates() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/typed16_for_column.c");
    run_native(&path);

    vybe_language_c::register();
    let mut bundle = vybe_compiler::projects::load(&path).unwrap();
    assert!(
        word_transfers(&bundle.prepared_module().unwrap()) > 0,
        "ordinary typed loop must exercise the optimized transfer"
    );
    validate_and_run(bundle.compile().unwrap());
    // Address exposure disables the loop rewrite without changing its result.
    assert!(bundle.sources[0].code.contains("int idx = 0;"));
    bundle.sources[0].code = bundle.sources[0].code.replace(
        "int idx = 0;",
        "int idx = 0; int *expose = &idx; (void)expose;",
    );
    assert_eq!(
        word_transfers(&bundle.prepared_module().unwrap()),
        0,
        "address-exposed index must retain the scalar loop"
    );
    validate_and_run(bundle.compile().unwrap());
}

fn run_native(path: &Path) {
    let executable = std::env::temp_dir().join(format!(
        "{}-native-{}",
        path.file_stem().unwrap().to_str().unwrap(),
        std::process::id()
    ));
    let status = std::process::Command::new("cc")
        .args(["-std=c11", "-O0", "-Wall", "-Wextra"])
        .arg(&path)
        .arg("-o")
        .arg(&executable)
        .status()
        .unwrap();
    assert!(status.success());
    let native = std::process::Command::new(&executable).status().unwrap();
    std::fs::remove_file(&executable).unwrap();
    assert!(native.success());
}

fn word_transfers(module: &vybe_ast::Module) -> usize {
    let mut count = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |expression| {
            if let vybe_ast::ExprKind::Call { callee, .. } = &expression.kind {
                count += usize::from(matches!(&callee.kind,
                    vybe_ast::ExprKind::Ident(name) if name == "__c_ptr_copy_column16"));
            }
        });
    }
    count
}

#[test]
fn pointer_arithmetic_chains_preserve_linear_and_managed_pointee_stride() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/pointer_arithmetic_chains.c");
    run_native(&path);
    vybe_language_c::register();
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    validate_and_run(bundle.compile().unwrap());
}

fn validate_and_run(chunks: Vec<vybe_runtime::Chunk>) {
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
    let mut abstract_imports = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(section) => {
                for import in section.into_imports() {
                    let import = import.unwrap();
                    if matches!(
                        import.ty,
                        wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)
                    ) {
                        abstract_imports.push(
                            import.module == "ecma:value"
                                && matches!(import.name, "abstractEq" | "abstractNe"),
                        );
                    }
                }
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                let mut operators = body.get_operators_reader().unwrap();
                while !operators.eof() {
                    if let wasmparser::Operator::Call { function_index }
                    | wasmparser::Operator::ReturnCall { function_index }
                    | wasmparser::Operator::RefFunc { function_index } =
                        operators.read().unwrap()
                    {
                        assert_ne!(
                            abstract_imports.get(function_index as usize),
                            Some(&true),
                            "emitted WASM references abstract equality at function {function_index}"
                        );
                    }
                }
            }
            _ => {}
        }
    }
    for chunk in &chunks {
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(
            !chunk.name.contains("dyneq") && !dump.contains("dyneq"),
            "{}: {dump}",
            chunk.name
        );
        let imports = if chunk.imports.is_empty() {
            &chunks[0].imports
        } else {
            &chunk.imports
        };
        let mut offset = 0;
        while offset < chunk.code.len() {
            let op = vybe_runtime::Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2))
                .unwrap();
            let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
            if op == vybe_runtime::Op::CALL {
                if let Some(import) = imports.get(chunk.read_u16(offset + 4) as usize) {
                    assert!(
                        !(import.module == "ecma:value"
                            && matches!(import.name.as_str(), "abstractEq" | "abstractNe")),
                        "{}@{offset}: {}:{}",
                        chunk.name,
                        import.module,
                        import.name
                    );
                }
            }
            assert!(next > offset && next <= chunk.code.len());
            offset = next;
        }
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    if let Err(error) = vm.run(chunks) {
        for frame in &error.call_stack {
            if let Some(chunk) = vm
                .chunks
                .iter()
                .find(|chunk| chunk.name == frame.chunk_name)
            {
                let mut offset = 0;
                while offset < chunk.code.len() && offset <= frame.offset + 16 {
                    let (instruction, next) =
                        vybe_runtime::debug::disassemble_instruction(chunk, offset);
                    if offset >= frame.offset.saturating_sub(100) {
                        eprintln!("{}@{offset}: {instruction}", chunk.name);
                    }
                    assert!(next > offset);
                    offset = next;
                }
            }
        }
        panic!("{error}");
    }
    assert_eq!(vm.pending_exit_code, 0);
}
