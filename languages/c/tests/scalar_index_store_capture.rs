use std::path::Path;
use vybe_ast::{BindingPattern, ExprKind, Module, StmtKind};

fn capture_declarations(module: &Module, function: &str) -> usize {
    let body = module
        .body
        .iter()
        .find_map(|statement| match &statement.kind {
            StmtKind::FunctionDecl { name, body, .. } if name == function => Some(body),
            _ => None,
        })
        .unwrap();
    body.iter()
        .filter_map(|statement| match &statement.kind {
            StmtKind::VarDecl { declarations, .. } => Some(declarations),
            _ => None,
        })
        .flatten()
        .filter(|declaration| {
            matches!(&declaration.pattern,
        BindingPattern::Ident(name) if name.starts_with("__c_index_scalar_"))
        })
        .count()
}

#[test]
fn scalar_index_rhs_has_one_ast_producer_and_invocation_local_captures() {
    let module = vybe_language_c::parse(
        "int index_once(void); int rhs_once(int); \
         int store(int *dest) { return dest[index_once()] = rhs_once(7); }",
    )
    .unwrap();
    assert_eq!(capture_declarations(&module, "store"), 3);
    let mut indices = 0;
    let mut producers = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |value| {
            if let ExprKind::Call { callee, .. } = &value.kind {
                if let ExprKind::Ident(name) = &callee.kind {
                    indices += usize::from(name == "index_once");
                    producers += usize::from(name == "rhs_once");
                }
            }
        });
    }
    assert_eq!(indices, 1);
    assert_eq!(producers, 1);
}

#[test]
fn scalar_index_capture_rejects_unproved_wide_pointer_and_volatile_values() {
    for (declarations, function) in [
        (
            "long long wide(void);",
            "int f(int *p) { return p[0] = wide(); }",
        ),
        (
            "double real(void);",
            "int f(int *p) { return p[0] = real(); }",
        ),
        ("", "long long f(long long *p) { return p[0] = 3; }"),
        ("", "int *f(int **p, int *q) { return p[0] = q; }"),
        ("", "int f(volatile int *p) { return p[0] = 3; }"),
    ] {
        let module = vybe_language_c::parse(&format!("{declarations} {function}")).unwrap();
        assert_eq!(capture_declarations(&module, "f"), 0, "{function}");
    }
}

fn word_transfers(module: &Module) -> usize {
    let mut count = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |value| {
            if let ExprKind::Call { callee, .. } = &value.kind {
                count += usize::from(matches!(&callee.kind,
                    ExprKind::Ident(name) if name == "__c_ptr_copy_column16"));
            }
        });
    }
    count
}

#[test]
fn affine_word_captures_require_ordinary_unexposed_loop_private_scratch() {
    for (global, local, tail, expected) in [
        ("", "short scratch = 0;", "return 0;", 1),
        ("", "short scratch = 0;", "return scratch;", 0),
        ("", "volatile short scratch = 0;", "return 0;", 0),
        (
            "",
            "short scratch = 0; short *expose = &scratch; (void)expose;",
            "return 0;",
            0,
        ),
        ("short scratch;", "", "return 0;", 0),
    ] {
        let source = format!(
            "{global} int scatter(short *d, short *s, int n, int width) {{ \
             int idx=0; int j; {local} \
             for(j=n;j;j--) {{ d[idx]=(scratch=*(s++)); idx+=width; }} {tail} }}"
        );
        let module = vybe_language_c::parse(&source).unwrap();
        assert_eq!(word_transfers(&module), expected, "{source}");
        // Both the no-touch zero-count arm and the complete-footprint check
        // remain required when private captures make a loop eligible again.
        if expected == 1 {
            let mut memory_bounds = 0;
            let mut zero_count = 0;
            for statement in &module.body {
                statement.walk_exprs(&mut |value| match &value.kind {
                    ExprKind::Call { callee, .. } => {
                        memory_bounds += usize::from(matches!(&callee.kind,
                            ExprKind::Ident(name) if name == "__c_ptr_memory_size"));
                    }
                    ExprKind::Binary {
                        op: vybe_ast::BinOp::Integer(vybe_ast::IntOp::Eq, _),
                        left,
                        right,
                    } if matches!(&left.kind, ExprKind::Ident(name) if name == "j")
                        && matches!(&right.kind, ExprKind::Lit(vybe_ast::Literal::Int(0))) =>
                    {
                        zero_count += 1;
                    }
                    _ => {}
                });
            }
            assert!(memory_bounds >= 3);
            assert!(zero_count >= 1);
        }
    }
}

#[test]
fn scalar_index_capture_matches_native_and_valid_static_wasm() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/scalar_index_store_capture.c");
    run_native(&path);
    vybe_language_c::register();
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let module = bundle.prepared_module().unwrap();
    for function in [
        "store_i16",
        "store_u16",
        "store_i32",
        "store_u32",
        "store_redirected",
        "recursive_store",
    ] {
        assert_eq!(capture_declarations(&module, function), 3, "{function}");
    }
    validate_and_run(bundle.compile().unwrap());
}

fn run_native(path: &Path) {
    let executable = std::env::temp_dir().join(format!(
        "{}-native-{}",
        path.file_stem().unwrap().to_str().unwrap(),
        std::process::id()
    ));
    assert!(
        std::process::Command::new("cc")
            .args([
                "-std=c11",
                "-O0",
                "-Wall",
                "-Wextra",
                "-fsanitize=address,undefined"
            ])
            .arg(&path)
            .arg("-o")
            .arg(&executable)
            .status()
            .unwrap()
            .success()
    );
    let native = std::process::Command::new(&executable).status().unwrap();
    std::fs::remove_file(executable).unwrap();
    assert!(native.success());
}

fn validate_and_run(chunks: Vec<vybe_runtime::Chunk>) -> vybe_runtime::Value {
    for chunk in &chunks {
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(
            !chunk.name.contains("dyneq") && !dump.contains("dyneq"),
            "{}",
            chunk.name
        );
        assert!(
            !dump.lines().any(
                |line| (line.contains("global.get") || line.contains("global.set"))
                    && (line.contains("__c_index_scalar_")
                        || chunk.name == "check"
                            && ["__c_struct_field_write", "__c_field_write_source", "__c_field_write_backing"]
                                .iter().any(|prefix| line.contains(prefix)))
            ),
            "capture escaped frame: {}",
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
                    if let wasmparser::Operator::Call { function_index }
                    | wasmparser::Operator::ReturnCall { function_index }
                    | wasmparser::Operator::RefFunc { function_index } =
                        operators.read().unwrap()
                    {
                        assert_ne!(abstract_imports.get(function_index as usize), Some(&true));
                    }
                }
            }
            _ => {}
        }
    }
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.memory.grow(1);
    let result = vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    result
}

#[test]
fn affine_word_expression_initializer_matches_native_once_with_mutation_and_overlap() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/affine_word_expression_initializer.c");
    run_native(&path);
    vybe_language_c::register();
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let module = bundle.prepared_module().unwrap();
    for (name, expected) in [
        ("scatter", 1),
        ("zero_does_not_touch_null_pointers", 0),
        ("zero_parameters_do_not_touch_null_pointers", 1),
    ] {
        let mut selected = module.clone();
        selected.body.retain(|statement| {
            matches!(&statement.kind,
            StmtKind::FunctionDecl { name: found, .. } if found == name)
        });
        assert_eq!(selected.body.len(), 1, "missing {name}");
        // The constant-NULL local currently retains an untyped Deref node;
        // do not certify it as a word load merely to select a bulk path.
        assert_eq!(word_transfers(&selected), expected, "{name}");
    }
    assert_eq!(word_transfers(&module), 2);
    let mut calls = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |value| {
            if let ExprKind::Call { callee, .. } = &value.kind {
                calls += usize::from(matches!(&callee.kind,
                    ExprKind::Ident(name) if name == "count_once"));
            }
        });
    }
    assert_eq!(
        calls, 3,
        "each unchanged initializer stays outside both loop versions"
    );
    validate_and_run(bundle.compile().unwrap());
}

fn find_word_guard(body: &[vybe_ast::Statement]) -> Option<vybe_ast::Expression> {
    for statement in body {
        match &statement.kind {
            StmtKind::If {
                cond,
                then_body,
                else_body: Some(other),
                ..
            } if matches!(
                other.as_slice(),
                [vybe_ast::Statement {
                    kind: StmtKind::For { .. },
                    ..
                }]
            ) =>
            {
                let mut has_transfer = false;
                for statement in then_body {
                    statement.walk_exprs(&mut |value| {
                        has_transfer |= matches!(&value.kind, ExprKind::Call { callee, .. }
                            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_ptr_copy_column16"));
                    });
                }
                if has_transfer {
                    return Some(cond.clone());
                }
            }
            StmtKind::Block(body) => {
                if let Some(guard) = find_word_guard(body) {
                    return Some(guard);
                }
            }
            _ => {}
        }
    }
    None
}

#[test]
fn expression_initializer_keeps_negative_count_and_full_footprint_guards() {
    vybe_language_c::register();
    let module = vybe_language_c::parse(
        "void scatter(short *d, short *s, int n, int width) { \
         int idx=0; int j; for(j=n+1-1;j;j--) { d[idx]=*(s++); idx+=width; } }",
    )
    .unwrap();
    assert_eq!(word_transfers(&module), 1);
    let body = module
        .body
        .iter()
        .find_map(|statement| match &statement.kind {
            StmtKind::FunctionDecl { name, body, .. } if name == "scatter" => Some(body),
            _ => None,
        })
        .unwrap();
    let guard = find_word_guard(body).expect("numeric bulk guard with scalar fallback");
    let profile = vybe_compiler::profile::parse_profile(vybe_language_c::profile_source()).unwrap();
    for (d, s, j, idx, width, expected) in [
        (1024, 2048, 3, 0, 2, 1),
        (0, 0, 0, 0, 2, 1),
        (1024, 2048, -1, 0, 2, 0),
        (1024, 2048, 2, i32::MAX as i64, 1, 0),
        (1024, 65535, 1, 0, 1, 0),
    ] {
        let mut check = module.clone();
        check.body = vec![
            vybe_ast::Statement::new(StmtKind::VarDecl {
                kind: vybe_ast::VarDeclKind::Let,
                declarations: [("d", d), ("s", s), ("j", j), ("idx", idx), ("width", width)]
                    .into_iter()
                    .map(|(name, value)| vybe_ast::VarDeclarator {
                        pattern: BindingPattern::Ident(name.into()),
                        type_hint: Some("int".into()),
                        init: Some(vybe_ast::Expression::int(value)),
                        array_bounds: None,
                        with_events: false,
                    })
                    .collect(),
            }),
            vybe_ast::Statement::new(StmtKind::Return(Some(guard.clone()))),
        ];
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile.clone())
            .compile(&check)
            .unwrap();
        assert_eq!(
            validate_and_run(chunks).as_i32(),
            expected,
            "d={d} s={s} j={j} idx={idx} width={width}"
        );
    }
}

#[test]
fn record_pointer_field_owner_addresses_match_write_through_and_release() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/record_pointer_field_owner.c");
    run_native(&path);
    vybe_language_c::register();
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    validate_and_run(bundle.compile().unwrap());
}

#[test]
fn record_w64_switch_transport_matches_native_without_module_policy() {
    for fixture in [
        "record_w64_switch_transport.c",
        "static_integral_switch.c",
        "unsigned_record_switch_selector.c",
    ] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../tests/c/doom_startup")
            .join(fixture);
        run_native(&path);
        vybe_language_c::register();
        let bundle = vybe_compiler::projects::load(&path).unwrap();
        assert_ne!(
            bundle
                .prepared_module()
                .unwrap()
                .directives
                .wide_integer_transport,
            Some(vybe_ast::WideIntegerTransport::Exact)
        );
        validate_and_run(bundle.compile().unwrap());
    }
}

#[test]
fn inline_record_array_store_keeps_owner_backing_index_stride_and_single_evaluation() {
    for fixture in [
        "inline_record_array_store.c",
        "scoped_embedded_patch_pointer.c",
    ] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../tests/c/doom_startup")
            .join(fixture);
        run_native(&path);
        vybe_language_c::register();
        let bundle = vybe_compiler::projects::load(&path).unwrap();
        validate_and_run(bundle.compile().unwrap());
    }
}

#[test]
fn inline_record_byte_copy_preserves_indexed_backing_and_evaluation() {
    for fixture in ["inline_record_byte_copy.c", "heap_record_name_comparison.c"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../tests/c/doom_startup")
            .join(fixture);
        run_native(&path);
        vybe_language_c::register();
        let bundle = vybe_compiler::projects::load(&path).unwrap();
        validate_and_run(bundle.compile().unwrap());
    }
}

#[test]
fn memcmp_scalar_arrays_use_object_bytes_without_changing_aliases() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/memcmp_scalar_array_layout.c");
    run_native(&path);
    vybe_language_c::register();
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    validate_and_run(bundle.compile().unwrap());
}

#[test]
fn enum_extents_require_source_order_constant_proof_before_global_storage() {
    let module = vybe_language_c::parse(
        "enum { WIDTH=16, HEIGHT=20, COLUMNS=WIDTH/2, WORDS=COLUMNS*HEIGHT }; \
         static short expected[WORDS]; static int expected_y[COLUMNS]; \
         int pixel_bytes=sizeof(expected); int cursor_bytes=sizeof(expected_y);",
    )
    .unwrap();
    let mut lengths: Vec<_> = module
        .body
        .iter()
        .filter_map(|statement| match &statement.kind {
            StmtKind::DataSegment { bytes, .. } => Some(bytes.len()),
            _ => None,
        })
        .collect();
    lengths.sort_unstable();
    assert_eq!(lengths, [32, 320]);
    for (name, bytes) in [("pixel_bytes", 320), ("cursor_bytes", 32)] {
        let init = module
            .body
            .iter()
            .find_map(|statement| {
                let StmtKind::VarDecl { declarations, .. } = &statement.kind else {
                    return None;
                };
                declarations.iter().find_map(|declaration| {
                matches!(&declaration.pattern, BindingPattern::Ident(binding) if binding == name)
                    .then_some(declaration.init.as_ref()).flatten()
            })
            })
            .unwrap();
        assert!(
            matches!(&init.kind, ExprKind::Lit(vybe_ast::Literal::Int(value)) if *value == bytes),
            "fixed linear object sizeof must be literal bytes, not numeric-address .length: {name}: {init:?}"
        );
    }
    for source in [
        "extern short a[4]; extern int b[8];",
        "enum { COUNT=4 }; extern short a[COUNT];",
        "_Thread_local short a[4];",
        "__thread int a[4];",
        "short a[LATER]; enum { LATER=4 };",
        "enum { BAD=missing, NEXT }; short a[NEXT];",
        "enum { BAD=1/0, NEXT }; short a[NEXT];",
    ] {
        let module = vybe_language_c::parse(source).unwrap();
        assert!(
            !module
                .body
                .iter()
                .any(|statement| matches!(&statement.kind, StmtKind::DataSegment { .. })),
            "unknown/invalid extent must not authorize linear storage: {source}"
        );
    }
}

#[test]
fn linear_global_scalar_arrays_preserve_aliases_offsets_and_definition_identity() {
    let module = vybe_language_c::parse("extern short a[]; short a[4]; short a[4]={1,2}; short a[4]; extern short a[4]; unsigned int b[3]; unsigned int c[3]; int d[3]={-1,2,3}; int memcmp(void*,void*,int); int f(short *p){return memcmp(a,p,8);}").unwrap();
    assert_eq!(
        module
            .body
            .iter()
            .filter(|statement| { matches!(&statement.kind, StmtKind::DataSegment { .. }) })
            .count(),
        4,
        "one range per object, not per declaration"
    );
    let mut snapshots = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |value| {
            if let ExprKind::Lambda { params, .. } = &value.kind {
                snapshots += usize::from(
                    params
                        .first()
                        .is_some_and(|param| param.name == "__c_memcmp_array"),
                );
            }
        });
    }
    assert_eq!(
        snapshots, 0,
        "linear global comparisons must not pack snapshots"
    );
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/linear_global_scalar_arrays.c");
    run_native(&path);
    vybe_language_c::register();
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    validate_and_run(bundle.compile().unwrap());
}

#[test]
fn memcmp_layout_adapter_requires_known_ordinary_scalar_array_and_builtin() {
    for (source, expected) in [
        (
            "int memcmp(void*,void*,int); int f(short *p) { static short a[4]; return memcmp(p,a,sizeof(a)); }",
            1,
        ),
        (
            "int memcmp(void*,void*,int); int f(int *p) { static int a[4]; return memcmp(a,p,sizeof(a)); }",
            1,
        ),
        (
            "int memcmp(void*,void*,int); volatile short a[4]; int f(short *p) { return memcmp(p,a,sizeof(a)); }",
            0,
        ),
        (
            "int memcmp(void*,void*,int); long long a[4]; int f(long long *p) { return memcmp(p,a,sizeof(a)); }",
            0,
        ),
        (
            "int memcmp(void*,void*,int); int f(short *p,short *q) { return memcmp(p,q,4); }",
            0,
        ),
        (
            "int memcmp(void*,void*,int); int f(short *p) { static short a[4]; return memcmp(a+1,p,4); }",
            0,
        ),
        (
            "short a[4]; int memcmp(void *p,void *q,int n) { return 19; } int f(short *p) { return memcmp(p,a,sizeof(a)); }",
            0,
        ),
    ] {
        let module = vybe_language_c::parse(source).unwrap();
        let mut snapshots = 0;
        for statement in &module.body {
            statement.walk_exprs(&mut |value| {
                if let ExprKind::Lambda { params, .. } = &value.kind {
                    snapshots += usize::from(
                        params
                            .first()
                            .is_some_and(|param| param.name == "__c_memcmp_array"),
                    );
                }
            });
        }
        assert_eq!(snapshots, expected, "{source}");
    }
}
