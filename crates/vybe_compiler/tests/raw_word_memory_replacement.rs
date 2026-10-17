use vybe_ast::{Argument, ExprKind, Expression, Lang, Module, Statement, StmtKind};
use vybe_runtime::{Chunk, registry};
use vybe_runtime::profile::{BuiltinDef, BuiltinEmit};
use registry::emitter_results::{self, EmitOutcome, EmitRequest, ResultRepresentation, TypedEmitterDef};
use vybe_compiler::primitives::argument_emission::{Definition, ProducedValueProof};
use std::sync::atomic::{AtomicUsize, Ordering};

static SPECIALIZATIONS: AtomicUsize = AtomicUsize::new(0);
static SPECIALIZERS: &[Definition] = &[Definition {
    operation: "pointers.linear_i32_store8", parser, producer: replacement,
    dispatch: argument_specializer,
}];

fn argument_specializer(name: &str, chunks: &mut Vec<Chunk>, current: usize,
                       argc: u8, line: u32, request: EmitRequest,
                       proofs: &[ProducedValueProof]) -> EmitOutcome {
    assert!(matches!(request, EmitRequest::Value | EmitRequest::DiscardedValue));
    assert_eq!(proofs, &[ProducedValueProof::Unknown, ProducedValueProof::Unknown]);
    SPECIALIZATIONS.fetch_add(1, Ordering::Relaxed);
    assert!(replacement(name, chunks, current, argc, line));
    EmitOutcome::Handled(ResultRepresentation::Unknown)
}

fn probe_dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, _: u8, line: u32) -> bool {
    if !matches!(name, "word_probe.read" | "raw_probe.read") { return false; }
    vybe_compiler::primitives::pointers::emit_linear_i32_load8_u(chunks, current, line);
    true
}

fn unknown_dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8,
                    line: u32, _: EmitRequest) -> EmitOutcome {
    if probe_dispatch(name, chunks, current, argc, line) {
        EmitOutcome::Handled(ResultRepresentation::Unknown)
    } else {
        EmitOutcome::Unhandled
    }
}

fn known_dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8,
                  line: u32, request: EmitRequest) -> EmitOutcome {
    if probe_dispatch(name, chunks, current, argc, line) {
        EmitOutcome::Handled(if request == EmitRequest::I32Word {
            ResultRepresentation::RawI32Word
        } else {
            ResultRepresentation::Unknown
        })
    } else {
        EmitOutcome::Unhandled
    }
}

fn replacement(name: &str, chunks: &mut Vec<Chunk>, current: usize, _: u8, line: u32) -> bool {
    match name {
        "pointers.linear_i32_load8_u" =>
            vybe_compiler::primitives::pointers::emit_linear_i32_load8_u(chunks, current, line),
        "pointers.linear_i32_store8" =>
            vybe_compiler::primitives::pointers::emit_linear_i32_store8(chunks, current, line),
        _ => return false,
    }
    true
}

fn parser(_: &str) -> Result<Module, String> { unreachable!() }
fn profile_source() -> &'static str { "[compiler]\n" }
fn register(name: &'static str, dispatch: registry::EmitDispatchFn) {
    registry::register_language(registry::LanguageDef {
        name, parse: parser, profile_source, emit_dispatch: Some(dispatch),
        normalize_class: None, register_tree: None, expand_source: None,
    });
}

fn check(read: &str, specialize: bool, discarded: bool) {
    let mut profile = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
    for (name, operation, arity) in [("fetch", read, 1),
        ("deposit", "pointers.linear_i32_store8", 2)] {
        profile.builtins.insert(name.into(), BuiltinDef {
            emit: BuiltinEmit::Common(operation.into()), discarded_emit: None,
            min_args: arity, max_args: arity, i32_condition: false, slot: None,
        });
    }
    let call = |name, values: Vec<Expression>| Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: values.into_iter().map(Argument::positional).collect(), optional: false,
    });
    let result = call("deposit", vec![Expression::int(4),
        call("fetch", vec![Expression::int(0)])]);
    let body = if discarded {
        vec![Statement::new(StmtKind::Expr(result)),
            Statement::new(StmtKind::Return(Some(Expression::int(0))))]
    } else {
        vec![Statement::new(StmtKind::Return(Some(result)))]
    };
    let module = Module {
        name: "<replacement-memory-word-proof>".into(), language: Lang::Unknown,
        body,
        imports: vec![], canon: Default::default(), directives: Default::default(),
    };
    let mut compiler = vybe_compiler::Compiler::with_profile(profile);
    if specialize {
        compiler = compiler.with_argument_emitters(parser, SPECIALIZERS);
    }
    let chunks = compiler.compile(&module).unwrap();
    let dump = chunks.iter().map(vybe_runtime::debug::disassemble).collect::<Vec<_>>().join("\n");
    assert!(dump.contains("i64.trunc_f64_s") && dump.contains("i32.wrap_i64"), "{dump}");
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata: false, include_name_section: true });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
}

// Namespace registration is irreversible; keep these cases in one isolated test binary.
#[test]
fn unknown_typed_producer_and_replacement_consumer_keep_legacy_store_abi() {
    register("word_probe", probe_dispatch);
    emitter_results::register(TypedEmitterDef {
        namespace: "word_probe", legacy_dispatch: probe_dispatch, dispatch: unknown_dispatch,
    }).unwrap();
    check("word_probe.read", false, false);
    register("raw_probe", probe_dispatch);
    emitter_results::register(TypedEmitterDef {
        namespace: "raw_probe", legacy_dispatch: probe_dispatch, dispatch: known_dispatch,
    }).unwrap();
    let mut proof = vec![Chunk::new("raw-producer")];
    assert_eq!(vybe_compiler::primitives::dispatch::emit_common_with_result(
        "raw_probe.read", &mut proof, 0, 1, 0, EmitRequest::I32Word),
        EmitOutcome::Handled(ResultRepresentation::RawI32Word));
    register("pointers", replacement);
    check("pointers.linear_i32_load8_u", false, false);
    for discarded in [false, true] {
        let before = SPECIALIZATIONS.load(Ordering::Relaxed);
        check("raw_probe.read", true, discarded);
        assert_eq!(SPECIALIZATIONS.load(Ordering::Relaxed), before + 1);
    }
}
