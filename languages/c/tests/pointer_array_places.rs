use std::path::Path;

#[test]
fn pointer_to_array_places_match_native_peg_and_standard_wasm() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    vybe_language_c::register();
    for name in [
        "tests/c/doom_startup/pointer_to_array_places.c",
        "tests/c/c_pointer_to_array_multidim/pointer_multidim_increment.c",
        "tests/c/c_pointer_to_array_multidim/pointer_multidim_deref.c",
        "tests/c/c_pointer_to_array_multidim/pointer_multidim_sizeof.c",
        "tests/c/c_pointer_to_array_multidim/pointer_multidim_3d.c",
        "tests/c/c_pointer_to_array_multidim/pointer_multidim_array_of_pointers.c",
        "tests/c/c_pointer_to_array_multidim/pointer_multidim_vla.c",
    ] {
        let path = root.join(name);
        run_native(&path);
        let mut asts = Vec::new();
        let mut modules = Vec::new();
        for parser in [
            vybe_language_c::parse_with_peg,
            vybe_language_c::parse_optimized,
        ] {
            let mut bundle = vybe_compiler::projects::load(&path).unwrap();
            bundle.language.parse = parser;
            asts.push(format!("{:?}", bundle.prepared_module().unwrap()));
            let chunks = bundle.compile().unwrap();
            let wasm = vybe_platform_wasm::write_wasm_with_options(
                &chunks,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false,
                    include_name_section: true,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&wasm)
                .unwrap();
            audit_equality(&wasm);
            let mut vm = vybe_runtime::VM::new();
            vybe_runtime::init_all_registered(
                &mut vm,
                &vybe_runtime::capabilities::Capabilities::all(),
            );
            vm.run(chunks)
                .unwrap_or_else(|error| panic!("{name}: {error}"));
            assert_eq!(vm.pending_exit_code, 0, "{name}");
            modules.push(wasm);
        }
        assert!(
            asts[0] == asts[1],
            "complete prepared PEG/native AST differs: {name}"
        );
        assert!(
            modules[0] == modules[1],
            "standard PEG/native WASM differs: {name}"
        );
    }
}

fn run_native(path: &Path) {
    let executable = std::env::temp_dir().join(format!(
        "pointer-array-native-{}-{}",
        std::process::id(),
        path.file_stem().unwrap().to_str().unwrap(),
    ));
    let status = std::process::Command::new("cc")
        .args(["-std=c11", "-O0", "-Wall", "-Wextra"])
        .arg(path)
        .arg("-o")
        .arg(&executable)
        .status()
        .unwrap();
    assert!(status.success(), "{}", path.display());
    let status = std::process::Command::new(&executable).status().unwrap();
    std::fs::remove_file(&executable).unwrap();
    assert!(status.success(), "{}", path.display());
}

fn audit_equality(wasm: &[u8]) {
    let mut abstract_imports = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(wasm) {
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
                            "abstract equality call/reference in emitted WASM"
                        );
                    }
                }
            }
            _ => {}
        }
    }
}
