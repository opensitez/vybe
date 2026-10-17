use std::sync::{Arc, Mutex};
use vybe_ast::{
    Argument, BinOp, BitLane, Directives, ExprKind, Expression, IntOp, Lang, Module,
    OperatorDispatch, Statement, StmtKind, UnaryOp,
};
use vybe_runtime::profile::{BuiltinDef, BuiltinEmit};
use vybe_runtime::{Chunk, Value, VM};

fn not(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(expr),
    })
}

fn predicate(name: &str) -> Expression {
    let call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: Vec::<Argument>::new(),
        optional: false,
    });
    Expression::new(ExprKind::Binary {
        op: BinOp::Integer(IntOp::Eq, BitLane::W32),
        left: Box::new(call),
        right: Box::new(Expression::int(1)),
    })
}

fn compile(expr: Expression, condition: bool, bitwise: bool) -> Vec<Chunk> {
    compile_with_dispatch(expr, condition, bitwise, OperatorDispatch::StaticBuiltin)
}

fn compile_with_dispatch(
    expr: Expression,
    condition: bool,
    bitwise: bool,
    dispatch: OperatorDispatch,
) -> Vec<Chunk> {
    let mut profile = vybe_compiler::profile::parse_profile(&format!(
        "[compiler]\nmaterialize_bool_results = true\nlogical_ops_bitwise_for_integers = {bitwise}\n"
    )).unwrap();
    for name in ["a", "b"] {
        profile.builtins.insert(
            name.into(),
            BuiltinDef {
                emit: BuiltinEmit::HostCall("test:condition".into(), name.into()),
                discarded_emit: None,
                min_args: 0,
                max_args: 0,
                i32_condition: false,
                slot: None,
            },
        );
    }
    let ret = |value| Statement::new(StmtKind::Return(Some(Expression::int(value))));
    let body = if condition {
        vec![Statement::new(StmtKind::If {
            cond: expr,
            then_body: vec![ret(10)],
            elifs: vec![],
            else_body: Some(vec![ret(20)]),
        })]
    } else {
        vec![Statement::new(StmtKind::Return(Some(expr)))]
    };
    let module = Module {
        name: "<static-negated-condition>".into(),
        language: Lang::Unknown,
        body,
        imports: vec![],
        canon: Default::default(),
        directives: Directives {
            operator_dispatch: Some(dispatch),
            logical_ops_bitwise_for_integers: Some(bitwise),
            ..Default::default()
        },
    };
    let chunks = vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap();
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
    chunks
}

fn run(chunks: Vec<Chunk>, a: i32, b: i32) -> (Value, Vec<String>) {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let calls = Arc::new(Mutex::new(Vec::new()));
    for (name, value) in [("a", a), ("b", b)] {
        let trace = calls.clone();
        vm.register_host_fn(
            "test:condition",
            name,
            Box::new(move |_, _| {
                trace.lock().unwrap().push(name.to_string());
                Value::F64(f64::from(value))
            }),
        );
    }
    let result = vm.run(chunks).unwrap();
    let trace = calls.lock().unwrap().clone();
    (result, trace)
}

#[test]
fn nested_static_negation_keeps_raw_conditions_and_evaluates_once() {
    for value in [0, 1] {
        for depth in 1..=3 {
            let mut expr = predicate("a");
            let mut expected = value == 1;
            for _ in 0..depth {
                expr = not(expr);
                expected = !expected;
            }
            let chunks = compile(expr, true, false);
            let dump = vybe_runtime::debug::disassemble(&chunks[0]);
            assert!(!dump.contains("wasm:js-boolean"), "{dump}");
            let (result, trace) = run(chunks, value, 0);
            assert_eq!(result.as_i32(), if expected { 10 } else { 20 });
            assert_eq!(trace, vec!["a"]);
        }
    }
}

#[test]
fn negated_short_circuit_preserves_skips_and_order() {
    for a in [0, 1] {
        for b in [0, 1] {
            for op in [BinOp::And, BinOp::Or] {
                let expr = not(Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(predicate("a")),
                    right: Box::new(predicate("b")),
                }));
                let expected = if op == BinOp::And {
                    !(a == 1 && b == 1)
                } else {
                    !(a == 1 || b == 1)
                };
                let skip = if op == BinOp::And { a == 0 } else { a == 1 };
                let (result, trace) = run(compile(expr, true, false), a, b);
                assert_eq!(result.as_i32(), if expected { 10 } else { 20 });
                assert_eq!(trace, if skip { vec!["a"] } else { vec!["a", "b"] });
            }
        }
    }
}

#[test]
fn value_negation_and_integer_bitwise_directive_keep_their_abis() {
    for value in [0, 1] {
        let (result, trace) = run(compile(not(predicate("a")), false, false), value, 0);
        assert_eq!(result, Value::Bool(value != 1));
        assert_eq!(trace, vec!["a"]);
    }
    for value in [0, -1] {
        let (result, trace) = run(compile(not(Expression::int(value)), true, true), 0, 0);
        assert_eq!(result.as_i32(), if value == 0 { 10 } else { 20 });
        assert!(trace.is_empty());
    }
}

#[test]
fn runtime_dispatch_keeps_fallback_and_matches_static_results() {
    for value in [0, 1] {
        let expr = not(predicate("a"));
        let static_chunks = compile(expr.clone(), true, false);
        let runtime_chunks =
            compile_with_dispatch(expr, true, false, OperatorDispatch::RuntimeProtocol);
        assert!(static_chunks[0].code.len() < runtime_chunks[0].code.len());
        let static_result = run(static_chunks, value, 0);
        let runtime_result = run(runtime_chunks, value, 0);
        assert_eq!(static_result, runtime_result);
    }
}
