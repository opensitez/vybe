use std::path::Path;
use std::time::Instant;

#[test]
fn scalar_w64_regions_reject_unproven_call_and_record_transport() {
    for source in [
        "unsigned long long unknown(void); int f(void) {
            unsigned long long word = unknown();
            switch(word) { case 0xffffffffffffffffULL: return 1; default: return 0; } }",
        "unsigned long long unknown(void); int f(void) {
            unsigned long long word = 0; word = (unsigned long long)unknown();
            switch(word) { case 0xffffffffffffffffULL: return 1; default: return 0; } }",
        "struct R { unsigned long long word; }; int f(struct R *record) {
            unsigned long long word = record->word;
            switch(word) { case 0xffffffffffffffffULL: return 1; default: return 0; } }",
        "int f(void) { volatile unsigned long long word = 0xffffffffffffffffULL;
            switch(word) { case 0xffffffffffffffffULL: return 1; default: return 0; } }",
    ] {
        let module = vybe_language_c::parse(source).unwrap();
        assert_eq!(module.directives.wide_integer_transport, None);
        let body = module
            .body
            .iter()
            .find_map(|statement| match &statement.kind {
                vybe_ast::StmtKind::FunctionDecl { name, body, .. } if name == "f" => Some(body),
                _ => None,
            })
            .unwrap();
        assert!(
            !body.iter().any(|statement| matches!(&statement.kind,
            vybe_ast::StmtKind::Directive { set, .. }
            if set.wide_integer_transport == Some(vybe_ast::WideIntegerTransport::Exact))),
            "unproven ABI/storage region selected: {source}"
        );
    }
}

#[test]
fn scalar_local_w64_switch_preserves_exact_values_without_wide_abi() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let source = root.join("tests/c/doom_startup/scalar_local_w64_switch.c");
    let native = std::env::temp_dir().join(format!("vybe-scalar-w64-{}", std::process::id()));
    let output = std::process::Command::new("cc")
        .args([
            "-std=c11",
            "-O0",
            "-fsanitize=address,undefined",
            "-fno-sanitize-recover=all",
        ])
        .arg(&source)
        .arg("-o")
        .arg(&native)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let native_start = Instant::now();
    let status = std::process::Command::new(&native).status().unwrap();
    let native_elapsed = native_start.elapsed();
    std::fs::remove_file(&native).unwrap();
    assert!(status.success());

    vybe_language_c::register();
    let compile_start = Instant::now();
    let project = vybe_compiler::projects::load(&source).unwrap();
    let prepared = project.prepared_module().unwrap();
    assert_eq!(prepared.directives.wide_integer_transport, None);
    let body = prepared
        .body
        .iter()
        .find_map(|statement| match &statement.kind {
            vybe_ast::StmtKind::FunctionDecl { name, body, .. } if name == "classify" => Some(body),
            _ => None,
        })
        .unwrap();
    assert!(
        matches!(&body[0].kind, vybe_ast::StmtKind::Directive { set, .. }
        if set.wide_integer_transport == Some(vybe_ast::WideIntegerTransport::Exact)),
        "scalar storage/switch proof did not select classify"
    );
    let chunks = project.compile().unwrap();
    let compile_elapsed = compile_start.elapsed();
    let classify = chunks
        .iter()
        .find(|chunk| chunk.name == "classify")
        .unwrap();
    let dump = vybe_runtime::debug::disassemble(classify);
    assert!(dump.matches("i64.eq").count() >= 4, "{dump}");
    for chunk in &chunks {
        assert!(
            !vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "{}",
            chunk.name
        );
    }
    let audit_start = Instant::now();
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: false,
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
                let mut reader = body.get_operators_reader().unwrap();
                while !reader.eof() {
                    match reader.read().unwrap() {
                        wasmparser::Operator::Call { function_index }
                        | wasmparser::Operator::ReturnCall { function_index }
                        | wasmparser::Operator::RefFunc { function_index } => {
                            assert_ne!(abstract_imports.get(function_index as usize), Some(&true));
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }
    let audit_elapsed = audit_start.elapsed();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let run_start = Instant::now();
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    eprintln!(
        "native execution={native_elapsed:?}; C compilation={compile_elapsed:?}; WASM write/validate/audit={audit_elapsed:?}; VM execution={:?}; WASM bytes={}",
        run_start.elapsed(),
        bytes.len()
    );
}
