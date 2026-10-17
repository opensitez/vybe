use std::collections::HashMap;
use std::path::Path;
use std::sync::atomic::Ordering;
use vybe_ast::{ExprKind, Module, Statement, StmtKind};
use vybe_runtime::{Chunk, DebugCommand, DebugEvent, DebugRequest, VM};

#[test]
fn actual_renderer_sources_activate_generic_backing_and_column_lowering() {
    vybe_language_c::register();
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../examples/chocolate-doom/src");
    let drawing = vybe_compiler::projects::load(&root.join("doom/r_draw.c"))
        .unwrap()
        .prepared_module()
        .unwrap();
    for name in ["R_DrawSpan", "R_DrawColumn"] {
        let mut body = drawing
            .body
            .iter()
            .find_map(|statement| match &statement.kind {
                StmtKind::FunctionDecl {
                    name: found, body, ..
                } if found == name => Some(body.clone()),
                _ => None,
            })
            .unwrap();
        assert_eq!(scalar_reference(&mut body), 1, "{name}");
    }
    let video = vybe_compiler::projects::load(&root.join("v_video.c"))
        .unwrap()
        .prepared_module()
        .unwrap();
    let patch = video.body.iter().find(|statement|
        matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name == "V_DrawPatch"))
        .unwrap();
    let mut columns = 0;
    patch.walk_exprs(&mut |value| {
        columns += usize::from(matches!(&value.kind, ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_ptr_copy_column")));
    });
    assert_eq!(columns, 1);
}

fn scalar_reference(body: &mut Vec<Statement>) -> usize {
    let mut count = 0;
    for statement in body {
        let mut restored = match &mut statement.kind {
            StmtKind::FunctionDecl { body, .. }
            | StmtKind::Block(body)
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. }
            | StmtKind::For { body, .. } => scalar_reference(body),
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut nested = scalar_reference(then_body);
                for (_, body) in elifs {
                    nested += scalar_reference(body);
                }
                if let Some(body) = else_body {
                    nested += scalar_reference(body);
                }
                nested
            }
            _ => 0,
        };
        if let StmtKind::If {
            then_body,
            else_body: Some(other),
            ..
        } = &statement.kind
        {
            let backing = then_body.len() == 1
                && other.len() == 1
                && matches!(
                    (&then_body[0].kind, &other[0].kind),
                    (StmtKind::DoWhile { .. }, StmtKind::DoWhile { .. })
                        | (StmtKind::While { .. }, StmtKind::While { .. })
                );
            let mut transfer = false;
            for fast in then_body {
                fast.walk_exprs(&mut |value| {
                    transfer |= matches!(&value.kind, ExprKind::Call { callee, .. }
                        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_ptr_copy_column"));
                });
            }
            let affine =
                transfer && other.len() == 1 && matches!(other[0].kind, StmtKind::While { .. });
            if backing || affine {
                *statement = other[0].clone();
                // Nested numeric and column versions represent the same loop.
                restored = 1;
            }
        }
        count += restored;
    }
    count
}

#[test]
fn affine_copy_preserves_cursor_results_managed_fallback_zero_count_and_overlap() {
    vybe_language_c::register();
    let path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/c/doom_startup/affine_byte_copy.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let optimized = bundle.prepared_module().unwrap();
    let mut baseline: Module = optimized.clone();
    assert_eq!(scalar_reference(&mut baseline.body), 3);
    let modules = HashMap::new();
    let scalar = bundle
        .compile_prepared_module(&baseline, &modules)
        .unwrap()
        .chunks;
    let normalized = bundle
        .compile_prepared_module(&optimized, &modules)
        .unwrap()
        .chunks;
    let scalar = run_counted(scalar).0;
    let normalized = run_counted(normalized).0;
    eprintln!("affine executed instructions: scalar={scalar}, normalized={normalized}");
    assert!(
        normalized < scalar,
        "scalar={scalar}, normalized={normalized}"
    );
}

fn run_counted(chunks: Vec<Chunk>) -> (u64, usize) {
    for chunk in &chunks {
        assert!(
            !vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "dynamic equality in renderer chunk {}",
            chunk.name
        );
        let imports = if chunk.imports.is_empty() {
            &chunks[0].imports
        } else {
            &chunk.imports
        };
        let mut offset = 0;
        while offset + 4 <= chunk.code.len() {
            let op = vybe_runtime::Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2))
                .expect("valid emitted opcode");
            if op == vybe_runtime::Op::CALL {
                if let Some(import) = imports.get(chunk.read_u16(offset + 4) as usize) {
                    assert!(
                        import.module != "ecma:value"
                            || !matches!(import.name.as_str(), "abstractEq" | "abstractNe"),
                        "dynamic equality host call in {} at {offset}",
                        chunk.name
                    );
                }
            }
            let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
            assert!(next > offset, "decoder did not advance in {}", chunk.name);
            offset = next;
        }
    }
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let (commands, rx) = std::sync::mpsc::channel();
    let (events, observations) = std::sync::mpsc::channel();
    vm.attach_debugger(rx, events, true);
    for command in [
        DebugCommand::StreamOpcodes { enabled: true },
        DebugCommand::Continue,
    ] {
        let (reply, _response) = std::sync::mpsc::channel();
        commands.send(DebugRequest { command, reply }).unwrap();
    }
    let report = vm.debug_report().unwrap();
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    let total = report
        .lock()
        .unwrap()
        .live
        .instructions
        .load(Ordering::Relaxed);
    let pixels = observations
        .try_iter()
        .filter(|event| matches!(event, DebugEvent::Opcode { chunk, .. } if chunk == "raster"))
        .count();
    (total, pixels)
}

#[test]
fn span_numeric_version_preserves_scalar_and_managed_paths_with_fewer_executed_opcodes() {
    compare_lookup_versions("linear_backing_span.c");
}

#[test]
fn column_numeric_version_preserves_stride_and_managed_lookup_with_fewer_executed_opcodes() {
    compare_lookup_versions("linear_backing_column.c");
}

#[test]
fn specialized_span_preserves_bytes_without_redundant_narrowing() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/linear_backing_span.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let optimized = bundle.prepared_module().unwrap();
    let mut baseline = optimized.clone();
    let mut restored = 0;
    for statement in &mut baseline.body {
        if !matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name == "raster") {
            continue;
        }
        statement.walk_exprs(&mut |value| {
            if matches!(&value.kind, ExprKind::Binary { left, right, .. }
                if matches!(&right.kind, ExprKind::Lit(vybe_ast::Literal::Int(255)))
                && matches!(&left.kind, ExprKind::Call { callee, .. }
                    if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_ptr_i32_load8_u")))
            {
                panic!("specialized unsigned byte load still has a redundant mask");
            }
        });
        statement.walk_exprs_mut(&mut |value| {
            if matches!(&value.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_ptr_i32_load8_u"))
            {
                let original = value.clone();
                value.kind = ExprKind::Binary {
                    op: vybe_ast::BinOp::Integer(vybe_ast::IntOp::And, vybe_ast::BitLane::W32),
                    left: Box::new(original),
                    right: Box::new(vybe_ast::Expression::int(255)),
                };
                restored += 1;
            }
        });
    }
    assert!(restored > 0, "no specialized byte loads found");
    let compile = |module: &Module| {
        let chunks = bundle
            .compile_prepared_module(module, &HashMap::new())
            .unwrap()
            .chunks;
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
    };
    let baseline = run_counted(compile(&baseline));
    let optimized = run_counted(compile(&optimized));
    eprintln!("span byte narrowing instructions: baseline={baseline:?}, optimized={optimized:?}");
    assert!(optimized.1 < baseline.1);
    assert!(optimized.0 + 64 < baseline.0);
}

fn compare_lookup_versions(fixture: &str) {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup")
        .join(fixture);
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let optimized = bundle.prepared_module().unwrap();
    let mut baseline: Module = optimized.clone();
    assert_eq!(scalar_reference(&mut baseline.body), 1);
    let modules = HashMap::new();
    let scalar = bundle
        .compile_prepared_module(&baseline, &modules)
        .unwrap()
        .chunks;
    let versioned = bundle
        .compile_prepared_module(&optimized, &modules)
        .unwrap()
        .chunks;
    for chunks in [&scalar, &versioned] {
        let bytes = vybe_platform_wasm::write_wasm_with_options(
            chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
    }
    let scalar = run_counted(scalar);
    let versioned = run_counted(versioned);
    eprintln!(
        "{fixture} executed instructions: scalar={scalar:?}, versioned={versioned:?} (total, raster)"
    );
    assert!(
        versioned.0 < scalar.0,
        "scalar={scalar:?}, versioned={versioned:?}"
    );
    assert!(
        versioned.1 < scalar.1,
        "scalar={scalar:?}, versioned={versioned:?}"
    );
}
