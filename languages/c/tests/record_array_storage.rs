use std::path::Path;

#[test]
fn packed_map_record_copy_preserves_initialized_bytes() {
    check("packed_map_record_copy_bytes", Some("records"));
}

#[test]
fn initialized_record_arrays_preserve_layout_aliases_and_single_evaluation() {
    check("initialized_record_array_storage", Some("globals"));
}

#[test]
fn unsigned_short_casts_narrow_before_record_byte_checks() {
    check("unsigned_short_cast_narrowing", None);
}

#[test]
fn indexed_record_inline_array_stores_preserve_typed_aliases() {
    check("indexed_record_inline_array_store", None);
}

fn check(fixture: &str, array_binding: Option<&str>) {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let source = root.join(format!("tests/c/doom_startup/{fixture}.c"));
    struct Artifact(std::path::PathBuf);
    impl Drop for Artifact {
        fn drop(&mut self) {
            let _ = std::fs::remove_file(&self.0);
        }
    }
    let executable = Artifact(std::env::temp_dir().join(format!(
        "vybe-record-array-{}-{fixture}",
        std::process::id(),
    )));
    let output = std::process::Command::new(std::env::var_os("CC").unwrap_or_else(|| "cc".into()))
        .args(["-std=gnu11", "-O0"])
        .arg("-idirafter")
        .arg(root.join("tools/compat"))
        .arg("-I")
        .arg(root.join("examples/chocolate-doom/src"))
        .arg(&source)
        .arg("-o")
        .arg(&executable.0)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(
        std::process::Command::new(&executable.0)
            .status()
            .unwrap()
            .success()
    );

    vybe_language_c::register();
    let project = vybe_compiler::projects::load(&source).unwrap();
    let prepared = project.prepared_module().unwrap();
    if let Some(name) = array_binding {
        let declaration = prepared.body.iter().find_map(|statement| {
        let vybe_ast::StmtKind::VarDecl { declarations, .. } = &statement.kind else { return None; };
        declarations.iter().find(|declaration|
            matches!(&declaration.pattern, vybe_ast::BindingPattern::Ident(binding) if binding == name))
    }).unwrap();
        assert!(declaration.array_bounds.is_none());
        let mut allocations = 0;
        let mut short_stores = 0;
        let mut byte_stores = 0;
        declaration.init.as_ref().unwrap().walk_exprs(&mut |value| {
        allocations += usize::from(matches!(&value.kind, vybe_ast::ExprKind::Call { callee, .. }
            if matches!(&callee.kind, vybe_ast::ExprKind::Ident(name) if name == "__c_ptr_linear_alloc")));
        short_stores += usize::from(matches!(&value.kind, vybe_ast::ExprKind::Call { callee, .. }
            if matches!(&callee.kind, vybe_ast::ExprKind::Ident(name) if name == "__c_ptr_i32_store16")));
        byte_stores += usize::from(matches!(&value.kind, vybe_ast::ExprKind::Call { callee, .. }
                if matches!(&callee.kind, vybe_ast::ExprKind::Ident(name)
                    if matches!(name.as_str(), "__c_ptr_i32_store8" | "__c_ptr_byte_store"))));
    });
        assert_eq!(
            allocations, 1,
            "{fixture}: initialized record array remained managed"
        );
        if fixture == "packed_map_record_copy_bytes" {
            assert_eq!(
                short_stores, 14,
                "two seven-short records, including sidenum[2]"
            );
        } else {
            assert_eq!(
                byte_stores, 6,
                "two initialized three-byte arrays must select typed byte stores"
            );
        }
    }
    let chunks = project.compile().unwrap();
    for chunk in &chunks {
        assert!(
            !vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "{}",
            chunk.name
        );
    }
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
                let mut operators = body.get_operators_reader().unwrap();
                while !operators.eof() {
                    match operators.read().unwrap() {
                        wasmparser::Operator::Call { function_index }
                        | wasmparser::Operator::ReturnCall { function_index }
                        | wasmparser::Operator::RefFunc { function_index } => {
                            assert_ne!(
                                abstract_imports.get(function_index as usize),
                                Some(&true),
                                "{fixture}: abstract equality call/reference {function_index}"
                            );
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
