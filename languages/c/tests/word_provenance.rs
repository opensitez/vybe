use std::collections::HashMap;
use std::path::Path;
use vybe_ast::{
    Argument, BinOp, BindingPattern, BitLane, ExprKind, Expression, Module, Statement, StmtKind,
    StorageAccess, TypeBinding,
};
use vybe_runtime::{Chunk, VM};

// Exercise the source pass against the same debug compiler and parser artifacts.
#[path = "../src/normalize_affine_bytes.rs"]
mod normalize_affine_bytes;
#[path = "../src/normalize_word_provenance.rs"]
mod provenance;
pub use vybe_language_c::{frontend, parse, parse_with_peg};

// Restore the unsigned Number exposure as a reference even when the linked
// parser already includes the source pass. This only changes word-only locals.
fn unsigned_reference(body: &mut [Statement]) -> usize {
    use std::collections::HashSet;
    fn expose(value: &mut Expression) {
        if matches!(&value.kind, ExprKind::Binary { op: BinOp::Integer(op, BitLane::W32), .. }
            if !op.is_comparison())
        {
            *value = Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(Expression::ident("__c_unsigned32_from_bits")),
                    args: vec![Argument::positional(value.clone())],
                    optional: false,
                },
                value.span,
            );
        }
    }
    fn declarations(body: &mut [Statement], names: &mut HashSet<String>) {
        for statement in body {
            if let StmtKind::VarDecl { declarations, .. } = &mut statement.kind {
                for declaration in declarations {
                    if let BindingPattern::Ident(name) = &declaration.pattern
                        && let Some(hint) = &mut declaration.type_hint
                        && hint.spelling() == "int32"
                        && hint.binding == TypeBinding::Checked
                        && hint.access == Some(StorageAccess::Ordinary)
                    {
                        hint.set_spelling("uint32");
                        names.insert(name.clone());
                        if let Some(value) = &mut declaration.init {
                            expose(value);
                        }
                    }
                }
            }
            normalize_affine_bytes::children(statement, &mut |body| declarations(body, names));
        }
    }
    let mut count = 0;
    for statement in body {
        if let StmtKind::FunctionDecl { body, .. } = &mut statement.kind {
            let mut names = HashSet::new();
            declarations(body, &mut names);
            count += names.len();
            for statement in body {
                statement.walk_exprs_mut(&mut |expression| {
                    if let ExprKind::Assign { target, value } = &mut expression.kind
                        && matches!(&target.kind, ExprKind::Ident(name) if names.contains(name))
                    {
                        expose(value);
                    }
                });
            }
        }
    }
    count
}

fn apply(source: &str) -> (usize, Module) {
    let mut module = parse(source).unwrap();
    let changed = provenance::normalize(&mut module.body);
    (changed, module)
}

#[test]
fn word_provenance_rejects_numeric_escapes_and_uncertain_storage() {
    for source in [
        "unsigned f(unsigned x) { unsigned a=x<<3; return a; }",
        "double f(unsigned x) { unsigned a=x<<3; return a+0.5; }",
        "long long f(unsigned x) { unsigned a=x<<3; return (long long)a+1; }",
        "void sink(unsigned); int f(unsigned x) { unsigned a=x<<3; sink(a); return a>>3; }",
        "int f(unsigned x) { volatile unsigned a=x<<3; return a>>3; }",
        "int f(unsigned x) { unsigned a=x<<3; unsigned *p=&a; return *p>>3; }",
        "int f(unsigned x) { unsigned a=x<<3; a=x; return a>>3; }",
        "unsigned f(unsigned x) { unsigned a=x<<3; return (a=x<<3); }",
        "void sink(unsigned); int f(unsigned x) { unsigned a=x<<3; sink(a=x<<3); return a>>3; }",
    ] {
        let (_, module) = apply(source);
        for statement in module.body {
            if let StmtKind::FunctionDecl { name, body, .. } = statement.kind
                && name == "f"
            {
                for statement in body {
                    if let StmtKind::VarDecl { declarations, .. } = statement.kind {
                        for declaration in declarations {
                            if matches!(&declaration.pattern, BindingPattern::Ident(name) if name == "a")
                            {
                                assert_ne!(
                                    declaration.type_hint.as_ref().map(|h| h.spelling()),
                                    Some("int32"),
                                    "{source}"
                                );
                            }
                        }
                    }
                }
            }
        }
    }
}

fn runtime() -> VM {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm
}

fn counted(chunks: Vec<Chunk>) -> (u64, HashMap<String, Vec<u64>>) {
    use std::sync::atomic::Ordering;
    use vybe_runtime::{DebugCommand, DebugEvent, DebugRequest};
    let mut vm = runtime();
    let (commands, rx) = std::sync::mpsc::channel();
    let (events, observations) = std::sync::mpsc::channel();
    let collector = std::thread::spawn(move || {
        let mut draws: HashMap<String, Vec<u64>> = HashMap::new();
        for event in observations {
            if let DebugEvent::Opcode { chunk, ip, .. } = event
                && matches!(chunk.as_str(), "R_DrawColumn" | "R_DrawSpan")
            {
                let calls = draws.entry(chunk).or_default();
                if ip == 0 {
                    calls.push(0);
                }
                if let Some(count) = calls.last_mut() {
                    *count += 1;
                }
            }
        }
        draws
    });
    vm.attach_debugger(rx, events, true);
    for command in [
        DebugCommand::StreamOpcodes { enabled: true },
        DebugCommand::Continue,
    ] {
        let (reply, _rx) = std::sync::mpsc::channel();
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
    drop(vm);
    (total, collector.join().unwrap())
}

fn validate(chunks: &[Chunk], name: &str) {
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
    if let Some(dir) = std::env::var_os("C_WORD_WASM_DIR") {
        std::fs::write(Path::new(&dir).join(format!("{name}.wasm")), bytes).unwrap();
    }
}

#[test]
#[cfg(debug_assertions)]
fn word_provenance_actual_draw_debug_measurements() {
    vybe_language_c::register();
    for fixture in ["doom_draw_kernel_heap", "doom_draw_kernel_edges"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../tests/c/doom_startup/{fixture}.c"));
        let bundle = vybe_compiler::projects::load(&path).unwrap();
        let mut baseline = bundle.prepared_module().unwrap();
        unsigned_reference(&mut baseline.body);
        let mut optimized = baseline.clone();
        let changed = provenance::normalize(&mut optimized.body);
        assert!(changed > 0, "no word-only locals in {fixture}");
        let compile = |module: &Module| {
            bundle
                .compile_prepared_module(module, &HashMap::new())
                .unwrap()
                .chunks
        };
        let before = compile(&baseline);
        let after = compile(&optimized);
        validate(&before, &format!("word-{fixture}-before"));
        validate(&after, &format!("word-{fixture}-after"));
        for chunks in [before, after] {
            let mut vm = runtime();
            vm.run(chunks).unwrap();
            assert_eq!(vm.pending_exit_code, 0);
        }
        // Full fixture parity ran above. Exclude its per-byte comparison work
        // from the traced measurement; the draw functions remain unchanged.
        let measured = |mut module: Module| {
            for statement in &mut module.body {
                if let StmtKind::FunctionDecl { name, body, .. } = &mut statement.kind
                    && name == "check_buffers"
                {
                    body.clear();
                }
            }
            counted(compile(&module))
        };
        let before = measured(baseline);
        let after = measured(optimized.clone());
        eprintln!(
            "{fixture}: bindings={changed} total={}->{}",
            before.0, after.0
        );
        for name in ["R_DrawColumn", "R_DrawSpan"] {
            let summarize = |counts: &HashMap<String, Vec<u64>>| {
                counts.get(name).map(|calls| {
                    (
                        calls.len(),
                        calls.first().copied(),
                        calls.iter().sum::<u64>(),
                    )
                })
            };
            eprintln!(
                "{name} (calls, first, total): {:?} -> {:?}",
                summarize(&before.1),
                summarize(&after.1)
            );
        }
        if changed > 0 {
            assert!(after.0 < before.0);
        }
        let previous = format!("{:?}", optimized.body);
        assert_eq!(provenance::normalize(&mut optimized.body), 0);
        assert_eq!(previous, format!("{:?}", optimized.body));
    }
}

#[test]
fn word_provenance_high_bits_unsigned_order_and_wrap_parity() {
    vybe_language_c::register();
    let source = include_str!("../../../tests/c/doom_startup/typed_word_provenance_high_bits.c");
    let mut baseline = parse(source).unwrap();
    unsigned_reference(&mut baseline.body);
    let mut optimized = baseline.clone();
    provenance::normalize(&mut optimized.body);
    let profile = std::sync::Arc::new(
        vybe_compiler::profile::parse_profile(vybe_language_c::profile_source()).unwrap(),
    );
    for (index, module) in [&baseline, &optimized].into_iter().enumerate() {
        let chunks = vybe_compiler::primitives::Compiler::with_shared_profile(profile.clone())
            .compile_with_imports(module)
            .unwrap()
            .chunks;
        validate(&chunks, &format!("word-high-bits-{index}"));
        let mut vm = runtime();
        vm.run(chunks).unwrap();
        assert_eq!(vm.pending_exit_code, 0);
    }
}
