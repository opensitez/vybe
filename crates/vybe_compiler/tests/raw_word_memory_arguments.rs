use vybe_ast::{Argument, BinOp, BitLane, ExprKind, Expression, IntOp, Lang, Module, Statement, StmtKind};
use vybe_runtime::{Chunk, VM};
use vybe_runtime::profile::{BuiltinDef, BuiltinEmit, LanguageProfile};

fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(), optional: false,
    })
}

fn profile(read: &str, store: &str) -> LanguageProfile {
    let mut profile = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
    for (name, operation, arity) in [("fetch", read, 1), ("deposit", store, 2),
        ("seed", "pointers.linear_i32_store", 2)] {
        profile.builtins.insert(name.into(), BuiltinDef {
            emit: BuiltinEmit::Common(operation.into()), discarded_emit: None,
            min_args: arity, max_args: arity, i32_condition: false, slot: None,
        });
    }
    profile
}

fn compile(profile: LanguageProfile, seed: i64, extra: Vec<Statement>, result: Expression) -> Vec<Chunk> {
    let mut body = vec![Statement::new(StmtKind::Expr(call("seed",
        vec![Expression::int(0), Expression::int(seed)])))];
    body.extend(extra);
    body.push(Statement::new(StmtKind::Return(Some(result))));
    let module = Module {
        name: "<raw-word-memory-arguments>".into(), language: Lang::Unknown,
        body, imports: vec![], canon: Default::default(), directives: Default::default(),
    };
    vybe_compiler::Compiler::with_profile(profile).compile(&module).unwrap()
}

fn dump(chunks: &[Chunk]) -> String {
    chunks.iter().map(vybe_runtime::debug::disassemble).collect::<Vec<_>>().join("\n")
}

fn validate(chunks: &[Chunk]) {
    let bytes = vybe_platform_wasm::write_wasm_with_options(chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
}

fn vm() -> VM {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.memory.grow(1);
    vm
}

#[test]
fn raw_load_store_arguments_remove_only_proven_conversions_and_preserve_assignment_values() {
    for (read, store, seed, expected, bytes) in [
        ("pointers.linear_i32_load8_u", "pointers.linear_i32_store8", 255, 255, vec![255]),
        ("pointers.linear_i32_load8_s", "pointers.linear_i32_store8", 255, -1, vec![255]),
        ("pointers.linear_i32_load16_u", "pointers.linear_i32_store16", 65535, 65535, vec![255, 255]),
        ("pointers.linear_i32_load16_s", "pointers.linear_i32_store16", 65535, -1, vec![255, 255]),
        ("pointers.linear_i32_load", "pointers.linear_i32_store", 0xffff_ffff, -1, vec![255; 4]),
    ] {
        let loaded = call("fetch", vec![Expression::int(0)]);
        let fast = compile(profile(read, store), seed, vec![],
            call("deposit", vec![Expression::int(4), loaded.clone()]));
        let slow = compile(profile(read, store), seed, vec![],
            call("deposit", vec![Expression::int(4), Expression::new(ExprKind::Cast {
                expr: Box::new(loaded), type_name: "double".into(),
            })]));
        let fast_dump = dump(&fast);
        let slow_dump = dump(&slow);
        // The seed's number-backed store still requires its original conversion.
        assert_eq!(fast_dump.matches("i64.trunc_f64_s").count(), 1, "{fast_dump}");
        assert_eq!(fast_dump.matches("i32.wrap_i64").count(), 1, "{fast_dump}");
        assert!(slow_dump.matches("i64.trunc_f64_s").count() > 1, "{slow_dump}");
        validate(&fast);
        validate(&slow);
        let mut runtime = vm();
        assert_eq!(runtime.run(fast).unwrap().as_i32(), expected, "{read}");
        runtime.memory.with_buffer_mut(|memory| assert_eq!(&memory[4..4 + bytes.len()], &bytes[..]));
        let mut runtime = vm();
        assert_eq!(runtime.run(slow).unwrap().as_i32(), expected, "coerced {read}");
    }
}

#[test]
fn unknown_host_and_number_arguments_retain_store_normalization() {
    let mut host = profile("pointers.linear_i32_load8_u", "pointers.linear_i32_store8");
    host.builtins.get_mut("fetch").unwrap().emit = BuiltinEmit::HostCall("test:word".into(), "fetch".into());
    for (profile, rhs) in [
        (host, call("fetch", vec![Expression::int(0)])),
        (profile("pointers.linear_i32_load8_u", "pointers.linear_i32_store8"), Expression::int(0xffff_ffff)),
        (profile("pointers.linear_i32_load8_u", "pointers.linear_i32_store8"),
            Expression::new(ExprKind::Lit(vybe_ast::Literal::Float(12.75)))),
    ] {
        let chunks = compile(profile, 7, vec![], call("deposit", vec![Expression::int(4), rhs]));
        assert_eq!(dump(&chunks).matches("i64.trunc_f64_s").count(), 2);
        assert_eq!(dump(&chunks).matches("i32.wrap_i64").count(), 2);
        validate(&chunks);
    }
}

#[test]
fn alias_adaptation_does_not_inherit_a_nested_word_report() {
    use vybe_ast::{InterfaceMember, Param, PassBy};
    let mut p = profile("pointers.linear_i32_load8_u", "pointers.linear_i32_store8");
    p.builtin_signatures.push(InterfaceMember::Method {
        name: "deposit".into(),
        params: [("address", PassBy::Value), ("value", PassBy::Alias)]
            .into_iter().map(|(name, pass_by)| Param {
                name: name.into(), type_hint: None, default: None, pass_by,
                is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false,
            }).collect(),
        return_type: None, is_sub: false, signature_source: None,
    });
    let chunks = compile(p, 255, vec![], call("deposit",
        vec![Expression::int(4), call("fetch", vec![Expression::int(0)])]));
    assert_eq!(dump(&chunks).matches("i64.trunc_f64_s").count(), 2);
    validate(&chunks);
    assert_eq!(vm().run(chunks).unwrap().as_i32(), 255);
}

#[test]
fn nested_loads_store_results_and_argument_effects_keep_order() {
    let p = profile("pointers.linear_i32_load", "pointers.linear_i32_store");
    let stored = call("deposit", vec![Expression::int(4),
        call("fetch", vec![call("fetch", vec![Expression::int(0)])])]);
    let chunks = compile(p.clone(), 12,
        vec![Statement::new(StmtKind::Expr(call("seed", vec![Expression::int(12), Expression::int(99)])))],
        Expression::new(ExprKind::Binary { op: BinOp::Integer(IntOp::And, BitLane::W32),
            left: Box::new(stored), right: Box::new(Expression::int(255)) }));
    validate(&chunks);
    assert_eq!(dump(&chunks).matches("i64.trunc_f64_s").count(), 2);
    assert_eq!(vm().run(chunks).unwrap().as_i32(), 99);

    let address = Expression::new(ExprKind::Sequence(vec![
        call("seed", vec![Expression::int(8), Expression::int(5)]), Expression::int(4),
    ]));
    let rhs = Expression::new(ExprKind::Sequence(vec![
        call("seed", vec![Expression::int(8), Expression::new(ExprKind::Binary {
            op: BinOp::Integer(IntOp::Add, BitLane::W32),
            left: Box::new(call("fetch", vec![Expression::int(8)])),
            right: Box::new(Expression::int(1)),
        })]), call("fetch", vec![Expression::int(0)]),
    ]));
    let chunks = compile(p, 13,
        vec![Statement::new(StmtKind::Expr(call("deposit", vec![address, rhs])))],
        call("fetch", vec![Expression::int(8)]));
    validate(&chunks);
    let mut runtime = vm();
    assert_eq!(runtime.run(chunks).unwrap().as_i32(), 6);
    runtime.memory.with_buffer_mut(|memory| assert_eq!(memory[4], 13));
}

#[test]
fn trapping_rhs_occurs_after_address_effect_and_before_store() {
    let address = Expression::new(ExprKind::Sequence(vec![
        call("seed", vec![Expression::int(8), Expression::int(5)]), Expression::int(4),
    ]));
    let chunks = compile(profile("pointers.linear_i32_load", "pointers.linear_i32_store"),
        13, vec![], call("deposit", vec![address, call("fetch", vec![Expression::int(0x7fff_ffff)])]));
    validate(&chunks);
    let mut runtime = vm();
    assert!(runtime.run(chunks).is_err());
    runtime.memory.with_buffer_mut(|memory| {
        assert_eq!(memory[8], 5);
        assert_eq!(memory[4], 0);
    });
}
