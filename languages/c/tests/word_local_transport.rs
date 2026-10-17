use std::collections::HashMap;
use std::path::Path;
use vybe_ast::{Directives, StmtKind, WordLocalTransport};
use vybe_compiler::word_local_proof::{WordInterpretation, prove, prove_subset};
use vybe_runtime::Op;

// Exercise the language-independent certification API in this focused cohort.
#[path = "../../../crates/vybe_compiler/tests/word_local_transport_proof.rs"]
mod generic_proof;

#[test]
fn actual_prepared_draw_span_selects_word_subset_and_records_boxed_conversion_baseline() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../examples/chocolate-doom/src/doom/r_draw.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let mut prepared = bundle.prepared_module().unwrap();
    prepared.directives.word_local_transport = Some(WordLocalTransport::Boxed);
    let body = prepared
        .body
        .iter()
        .find_map(|statement| match &statement.kind {
            StmtKind::FunctionDecl { name, body, .. } if name == "R_DrawSpan" => Some(body),
            _ => None,
        })
        .unwrap();
    let profile = vybe_compiler::profile::parse_profile(vybe_language_c::profile_source()).unwrap();
    let resolve_cast = |name: &str| {
        let width = profile.integer_cast_widths.get(name)?;
        (width.bits == Some(32)).then_some(if width.signed {
            WordInterpretation::Signed
        } else {
            WordInterpretation::Unsigned
        })
    };
    let resolve_binding = |hint: &vybe_ast::TypeHint| hint.word_value;
    let policy = Directives {
        word_local_transport: Some(WordLocalTransport::ProvenW32),
        ..prepared.directives.clone()
    };
    assert!(
        prove(&policy, body, resolve_binding).is_none(),
        "closed toy proof unexpectedly covered memory-bearing renderer"
    );
    assert!(
        prove_subset(&Directives::default(), body, resolve_binding, resolve_cast)
            .eligible
            .is_empty()
    );
    let subset = prove_subset(&policy, body, resolve_binding, resolve_cast);
    for name in ["position", "step"] {
        assert!(
            subset.eligible.iter().any(|binding| binding.name == name
                && binding.interpretation == WordInterpretation::Signed),
            "{name}: {subset:?}"
        );
    }
    assert!(!subset.eligible.iter().any(|binding| binding.name == "dest"));

    let baseline = bundle
        .compile_prepared_module(&prepared, &HashMap::new())
        .unwrap();
    // The bundle currently exposes chunks, not the new advisory sidecar.
    // No physical proofs are issued by the compiler constructor yet.
    for chunk in &baseline.chunks {
        assert!(
            !vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "dynamic equality in {}",
            chunk.name
        );
    }
    validate_no_equality_calls(&baseline.chunks);
    let chunk = baseline
        .chunks
        .iter()
        .find(|chunk| chunk.name == "R_DrawSpan")
        .unwrap();
    let mut truncations = 0;
    let mut wraps = 0;
    let mut offset = 0;
    while offset + 4 <= chunk.code.len() {
        let op = Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2)).unwrap();
        truncations += usize::from(op == Op::I64_TRUNC_F64_S);
        wraps += usize::from(op == Op::I32_WRAP_I64);
        let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
        assert!(next > offset);
        offset = next;
    }
    eprintln!(
        "actual R_DrawSpan boxed baseline: truncations={truncations}, wraps={wraps}; no physical proofs emitted"
    );
    assert!(truncations > 0 && wraps > 0);
    let mut optimized = prepared.clone();
    optimized.directives.word_local_transport = Some(WordLocalTransport::ProvenW32);
    let optimized = bundle
        .compile_prepared_module(&optimized, &HashMap::new())
        .unwrap();
    validate_no_equality_calls(&optimized.chunks);
    let chunk = optimized
        .chunks
        .iter()
        .find(|chunk| chunk.name == "R_DrawSpan")
        .unwrap();
    let (optimized_truncations, optimized_wraps) = narrowings(chunk);
    assert!(optimized_truncations < truncations);
    assert!(optimized_wraps < wraps);
    eprintln!(
        "R_DrawSpan narrowing pairs: {truncations}/{wraps} -> {optimized_truncations}/{optimized_wraps}"
    );
}

fn narrowings(chunk: &vybe_runtime::Chunk) -> (usize, usize) {
    let (mut truncations, mut wraps, mut offset) = (0, 0, 0);
    while offset + 4 <= chunk.code.len() {
        let op = Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2)).unwrap();
        truncations += usize::from(op == Op::I64_TRUNC_F64_S);
        wraps += usize::from(op == Op::I32_WRAP_I64);
        let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
        assert!(next > offset);
        offset = next;
    }
    (truncations, wraps)
}

#[test]
fn actual_draw_frames_reuse_boxed_word_operands_with_native_parity_and_instruction_drop() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let source = root.join("tests/c/doom_startup/doom_draw_word_reuse.c");
    let native = std::env::temp_dir().join(format!("vybe-word-draw-{}", std::process::id()));
    struct Artifact(std::path::PathBuf);
    impl Drop for Artifact {
        fn drop(&mut self) {
            let _ = std::fs::remove_file(&self.0);
        }
    }
    let native = Artifact(native);
    let output = std::process::Command::new(std::env::var_os("CC").unwrap_or_else(|| "cc".into()))
        .args([
            "-std=gnu11",
            "-O0",
            "-ffunction-sections",
            "-fdata-sections",
        ])
        .arg("-idirafter")
        .arg(root.join("tools/compat"))
        .arg("-I")
        .arg(root.join("examples/chocolate-doom/src"))
        .arg("-I")
        .arg(root.join("examples/chocolate-doom/src/doom"))
        .arg(&source)
        .arg(if cfg!(target_os = "macos") {
            "-Wl,-dead_strip"
        } else {
            "-Wl,--gc-sections"
        })
        .arg("-o")
        .arg(&native.0)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(
        std::process::Command::new(&native.0)
            .status()
            .unwrap()
            .success()
    );
    vybe_language_c::register();
    let bundle = vybe_compiler::projects::load(&source).unwrap();
    let mut baseline = bundle.prepared_module().unwrap();
    baseline.directives.word_local_transport = Some(WordLocalTransport::Boxed);
    let mut optimized = baseline.clone();
    optimized.directives.word_local_transport = Some(WordLocalTransport::ProvenW32);
    let baseline = bundle
        .compile_prepared_module(&baseline, &HashMap::new())
        .unwrap()
        .chunks;
    let optimized = bundle
        .compile_prepared_module(&optimized, &HashMap::new())
        .unwrap()
        .chunks;
    let mut wasm_sizes = Vec::with_capacity(2);
    for chunks in [&baseline, &optimized] {
        for chunk in chunks {
            assert!(
                !vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
                "{}",
                chunk.name
            );
        }
        wasm_sizes.push(validate_no_equality_calls(chunks));
    }
    let bytecode_sizes: Vec<usize> = [&baseline, &optimized]
        .into_iter()
        .map(|chunks| chunks.iter().map(|chunk| chunk.code.len()).sum())
        .collect();
    eprintln!(
        "actual draw bytecode/Wasm bytes: {}/{} -> {}/{}",
        bytecode_sizes[0], wasm_sizes[0], bytecode_sizes[1], wasm_sizes[1]
    );
    assert!(bytecode_sizes[1] < bytecode_sizes[0]);
    assert!(wasm_sizes[1] < wasm_sizes[0]);
    let count = |chunks: &[vybe_runtime::Chunk]| {
        chunks
            .iter()
            .filter(|chunk| matches!(chunk.name.as_str(), "R_DrawSpan" | "R_DrawColumn"))
            .map(|chunk| narrowings(chunk).0)
            .sum::<usize>()
    };
    assert!(count(&optimized) < count(&baseline));
    let before = run_draw_region(baseline);
    let after = run_draw_region(optimized);
    eprintln!(
        "actual four-frame draw region: {before} -> {after} instructions ({} removed)",
        before - after
    );
    assert!(
        after < before,
        "boxed proof did not reduce actual renderer work"
    );
}

fn run_draw_region(chunks: Vec<vybe_runtime::Chunk>) -> u64 {
    use std::sync::atomic::Ordering;
    use vybe_runtime::{DebugCommand, DebugEvent, DebugRequest, VM};
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let (commands, requests) = std::sync::mpsc::channel();
    let (events, observations) = std::sync::mpsc::channel();
    vm.attach_debugger(requests, events, true);
    let report = vm.debug_report().unwrap();
    for name in ["bench_draw_frames", "bench_draw_done"] {
        let (reply, _) = std::sync::mpsc::channel();
        commands
            .send(DebugRequest {
                command: DebugCommand::BreakFunction {
                    name: name.into(),
                    condition: None,
                },
                reply,
            })
            .unwrap();
    }
    let observer = std::thread::spawn(move || {
        let mut begin = None;
        let mut end = None;
        while let Ok(event) = observations.recv() {
            match event {
                DebugEvent::Paused { location, .. } => {
                    let count = report
                        .lock()
                        .unwrap()
                        .live
                        .instructions
                        .load(Ordering::Relaxed);
                    match location.chunk_name.as_str() {
                        "bench_draw_frames" => {
                            assert!(begin.replace(count).is_none());
                        }
                        "bench_draw_done" => {
                            assert!(end.replace(count).is_none());
                        }
                        _ => {}
                    }
                    let (reply, _) = std::sync::mpsc::channel();
                    commands
                        .send(DebugRequest {
                            command: DebugCommand::Continue,
                            reply,
                        })
                        .unwrap();
                }
                DebugEvent::Exited { .. } => break,
                _ => {}
            }
        }
        end.unwrap() - begin.unwrap()
    });
    let result = vm.run(chunks);
    let exit = vm.pending_exit_code;
    drop(vm);
    let instructions = observer.join().unwrap();
    result.unwrap();
    assert_eq!(exit, 0);
    instructions
}

fn validate_no_equality_calls(chunks: &[vybe_runtime::Chunk]) -> usize {
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        chunks,
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
                                "actual abstract equality call/reference {function_index}"
                            );
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }
    bytes.len()
}
