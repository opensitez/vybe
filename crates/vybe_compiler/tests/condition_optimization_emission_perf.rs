use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::{bundle, expressions, globals, platforms};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/condition_optimization_logical.rs"]
mod baseline;

#[derive(Clone, Debug)]
enum Condition {
    Leaf(i32, &'static str),
    And(Box<Self>, Box<Self>),
    Or(Box<Self>, Box<Self>),
}
fn leaf(value: i32, name: &'static str) -> Condition {
    Condition::Leaf(value, name)
}
fn and(left: Condition, right: Condition) -> Condition {
    Condition::And(Box::new(left), Box::new(right))
}
fn or(left: Condition, right: Condition) -> Condition {
    Condition::Or(Box::new(left), Box::new(right))
}
fn emit(chunk: &mut Chunk, condition: &Condition, old: bool) {
    match condition {
        Condition::Leaf(value, name) => {
            chunk.emit_string_const(&format!("{value}:{name}"), 0);
            core_wasm::i32_const(chunk, 0, 0);
            let char_at = chunk.add_import("ecma:string", "charAt");
            chunk.emit_call(char_at, 2, 0);
            chunk.emit_string_const("1", 0);
            let equals = chunk.add_import("wasm:js-string", "equals");
            chunk.emit_call(equals, 2, 0);
        }
        Condition::And(left, right) | Condition::Or(left, right) => {
            emit(chunk, left, old);
            let is_and = matches!(condition, Condition::And(..));
            let branch = match (old, is_and) {
                (true, true) => baseline::emit_and_i32_start(chunk, 0),
                (false, true) => expressions::emit_and_i32_start(chunk, 0),
                (true, false) => baseline::emit_or_i32_start(chunk, 0),
                (false, false) => expressions::emit_or_i32_start(chunk, 0),
            };
            emit(chunk, right, old);
            if old {
                baseline::emit_short_circuit_end(chunk, branch);
            } else if is_and {
                expressions::emit_and_i32_end(chunk, branch);
            } else {
                expressions::emit_short_circuit_end(chunk, branch);
            }
        }
    }
}
fn program(condition: &Condition, old: bool) -> Vec<Chunk> {
    let mut chunk = Chunk::new("<raw-logical-condition>");
    emit(&mut chunk, condition, old);
    chunk.emit_if_i32(0);
    core_wasm::i32_const(&mut chunk, 0, 10);
    chunk.emit_else(0);
    core_wasm::i32_const(&mut chunk, 0, 20);
    chunk.emit_end(0);
    chunk.emit_op(Op::RETURN, 0);
    let mut chunks = vec![chunk];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn runtime() -> (VM, Arc<Mutex<Vec<String>>>) {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let calls = Arc::new(Mutex::new(Vec::new()));
    let observed = calls.clone();
    // Instrument the existing charAt host call, preserving its actual result.
    // wasm string length has a VM intrinsic that bypasses host overrides.
    vm.register_host_fn(
        "ecma:string",
        "charAt",
        Box::new(move |_ctx, args| {
            let text = args[0].as_str();
            let (value, name) = text.split_once(':').unwrap();
            observed.lock().unwrap().push(name.into());
            assert_eq!(args[1], Value::I32(0));
            Value::String(value.into())
        }),
    );
    (vm, calls)
}
#[test]
fn raw_logical_conditions_preserve_order_skips_and_nested_results() {
    for a in [0, 1] {
        for b in [0, 1] {
            for c in [0, 1] {
                for (condition, expected, trace) in [
                    (
                        and(leaf(a, "a"), leaf(b, "b")),
                        a != 0 && b != 0,
                        if a == 0 { vec!["a"] } else { vec!["a", "b"] },
                    ),
                    (
                        or(leaf(a, "a"), leaf(b, "b")),
                        a != 0 || b != 0,
                        if a != 0 { vec!["a"] } else { vec!["a", "b"] },
                    ),
                    (
                        and(or(leaf(a, "a"), leaf(b, "b")), leaf(c, "c")),
                        (a != 0 || b != 0) && c != 0,
                        if a != 0 {
                            vec!["a", "c"]
                        } else if b != 0 {
                            vec!["a", "b", "c"]
                        } else {
                            vec!["a", "b"]
                        },
                    ),
                    (
                        or(and(leaf(a, "a"), leaf(b, "b")), leaf(c, "c")),
                        (a != 0 && b != 0) || c != 0,
                        if a == 0 {
                            vec!["a", "c"]
                        } else if b == 0 {
                            vec!["a", "b", "c"]
                        } else {
                            vec!["a", "b"]
                        },
                    ),
                ] {
                    for old in [true, false] {
                        let chunks = program(&condition, old);
                        let validation = wasmparser::Validator::new_with_features(
                            wasmparser::WasmFeatures::all(),
                        )
                        .validate_all(&vybe_platform_wasm::write_wasm(&chunks));
                        if old {
                            assert!(
                                validation.is_err(),
                                "captured boxed block result feeds i32 IF"
                            );
                        } else {
                            validation.unwrap();
                        }
                        let (mut vm, calls) = runtime();
                        let result = vm.run(chunks).unwrap();
                        assert_eq!(
                            result,
                            Value::I32(if expected { 10 } else { 20 }),
                            "old={old} a={a} b={b} c={c} condition={condition:?} actual_trace={:?} expected_trace={trace:?}",
                            *calls.lock().unwrap()
                        );
                        assert_eq!(
                            *calls.lock().unwrap(),
                            trace,
                            "old={old} a={a} b={b} c={c} condition={condition:?}"
                        );
                    }
                }
            }
        }
    }
}

#[test]
fn c_profile_ast_if_conditions_use_valid_raw_logical_results() {
    use vybe_ast::{
        BinOp, BindingPattern, Directives, ExprKind, Expression, Lang, Module, OperatorDispatch,
        Statement, StmtKind, TypeResolution, VarDeclKind, VarDeclarator,
    };
    fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
    }
    fn test(name: &str) -> Expression {
        Expression::ident(name)
    }
    let profile =
        vybe_compiler::profile::parse_profile(include_str!("../../../languages/c/src/profile"))
            .unwrap();
    for a in [false, true] {
        for b in [false, true] {
            for c in [false, true] {
                for (condition, expected) in [
                    (binary(BinOp::And, test("a"), test("b")), a && b),
                    (binary(BinOp::Or, test("a"), test("b")), a || b),
                    (
                        binary(
                            BinOp::And,
                            binary(BinOp::And, test("a"), test("b")),
                            test("c"),
                        ),
                        a && b && c,
                    ),
                    (
                        binary(
                            BinOp::Or,
                            binary(BinOp::Or, test("a"), test("b")),
                            test("c"),
                        ),
                        a || b || c,
                    ),
                ] {
                    let mut body = Vec::new();
                    for (name, value) in [("a", a), ("b", b), ("c", c)] {
                        body.push(Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(name.into()),
                                type_hint: None,
                                init: Some(Expression::bool(value)),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Var,
                        }));
                    }
                    let diagnostic = format!("a={a} b={b} c={c} condition={condition:?}");
                    body.push(Statement::new(StmtKind::If {
                        cond: condition,
                        then_body: vec![Statement::new(StmtKind::Return(Some(Expression::int(
                            10,
                        ))))],
                        elifs: vec![],
                        else_body: Some(vec![Statement::new(StmtKind::Return(Some(
                            Expression::int(20),
                        )))]),
                    }));
                    let module = Module {
                        name: "<c-logical-if>".into(),
                        language: Lang::Unknown,
                        body,
                        imports: vec![],
                        canon: Default::default(),
                        directives: Directives {
                            type_resolution: Some(TypeResolution::Static),
                            operator_dispatch: Some(OperatorDispatch::StaticBuiltin),
                            ..Default::default()
                        },
                    };
                    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile.clone())
                        .compile(&module)
                        .unwrap();
                    let bytes = vybe_platform_wasm::write_wasm(&chunks);
                    if let Err(error) =
                        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                            .validate_all(&bytes)
                    {
                        std::fs::write("/private/tmp/condition-c-ast.wasm", &bytes).unwrap();
                        panic!(
                            "{diagnostic}: {error}; diagnostic export /private/tmp/condition-c-ast.wasm"
                        );
                    }
                    let mut vm = VM::new();
                    platforms::register_platforms_all(&mut vm);
                    assert_eq!(
                        vm.run(chunks).unwrap().as_i32(),
                        if expected { 10 } else { 20 }
                    );
                }
            }
        }
    }
}

#[test]
fn c_integer_comparison_records_separate_signed_conversion_export_gap() {
    use vybe_ast::{
        BinOp, BindingPattern, Directives, ExprKind, Expression, Lang, Module, OperatorDispatch,
        Statement, StmtKind, TypeHint, TypeResolution, VarDeclKind, VarDeclarator,
    };
    // No AND/OR here: the boxed-f64 -> signed-i64 conversion fails independently
    // of short-circuit lowering. Preserve this diagnostic until its owning
    // exporter pass implements checked unboxing for signed conversions.
    let module = Module {
        name: "<c-signed-conversion-gap>".into(),
        language: Lang::Unknown,
        imports: vec![],
        canon: Default::default(),
        directives: Directives {
            type_resolution: Some(TypeResolution::Static),
            operator_dispatch: Some(OperatorDispatch::StaticBuiltin),
            ..Default::default()
        },
        body: vec![
            Statement::new(StmtKind::VarDecl {
                kind: VarDeclKind::Var,
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident("value".into()),
                    type_hint: Some(TypeHint::converting("int")),
                    init: Some(Expression::int(0)),
                    array_bounds: None,
                    with_events: false,
                }],
            }),
            Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(Expression::ident("value")),
                right: Box::new(Expression::int(0)),
            })))),
        ],
    };
    let profile =
        vybe_compiler::profile::parse_profile(include_str!("../../../languages/c/src/profile"))
            .unwrap();
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap();
    let validation = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(&chunks));
    match validation {
        Err(error) => assert!(
            error.to_string().contains("expected f64, found externref"),
            "unexpected gap: {error}"
        ),
        Ok(_) => panic!(
            "signed conversion export gap repaired; turn this diagnostic into a positive check"
        ),
    }
}
