use std::sync::{Arc, Mutex, Once};
use vybe_ast::{
    Argument, BinOp, Directives, ExprKind, Expression, LambdaBody, Lang, Module, OperatorDispatch,
    Statement, StmtKind, TypeResolution,
};
use vybe_runtime::profile::{BuiltinDef, BuiltinEmit, LanguageProfile};
use vybe_runtime::registry::emitter_results::{
    self, EmitOutcome, EmitRequest, ResultRepresentation, TypedEmitterDef,
};
use vybe_runtime::{Chunk, VM, Value, registry};

fn call(name: &str, values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: values.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn ret(value: Expression) -> Statement {
    Statement::new(StmtKind::Return(Some(value)))
}

fn profile() -> LanguageProfile {
    let mut profile = vybe_compiler::profile::parse_profile(
        "[compiler]\nfunction_references = true\nmaterialize_bool_results = true\n\
         [builtin_types]\nint = [\"int\", \"integer\"]\ndouble = [\"double\"]\n\
         [builtin_slots.int]\nadd = \"opcode:i32.add\"\nsub = \"opcode:i32.sub\"\n\
         mul = \"opcode:i32.mul\"\neq = \"opcode:i32.eq\"\nlt = \"opcode:i32.lt_s\"\n\
         [builtin_slots.double]\nadd = \"opcode:f64.add\"\n\
         [integer_cast_widths]\nint = { bits = 32, signed = true }\n",
    )
    .unwrap();
    for (name, emit, arity) in [
        (
            "fetch",
            BuiltinEmit::Common("pointers.linear_i32_load8_u".into()),
            1,
        ),
        (
            "seed",
            BuiltinEmit::Common("pointers.linear_i32_store".into()),
            2,
        ),
        (
            "left",
            BuiltinEmit::HostCall("test:static-word".into(), "left".into()),
            0,
        ),
        (
            "right",
            BuiltinEmit::HostCall("test:static-word".into(), "right".into()),
            0,
        ),
        (
            "identity",
            BuiltinEmit::HostCall("test:static-word".into(), "identity".into()),
            1,
        ),
        (
            "replace",
            BuiltinEmit::HostCall("test:static-word".into(), "replace".into()),
            0,
        ),
    ] {
        profile.builtins.insert(
            name.into(),
            BuiltinDef {
                emit,
                discarded_emit: None,
                min_args: arity,
                max_args: arity,
                i32_condition: false,
                slot: None,
            },
        );
        // This metadata selects a static operator, never a physical result proof.
        profile
            .builtin_return_types
            .insert(name.into(), "integer".into());
    }
    profile
}

fn compile(profile: LanguageProfile, body: Vec<Statement>) -> Vec<Chunk> {
    compile_case(profile, body, "default")
}

fn compile_case(profile: LanguageProfile, body: Vec<Statement>, case: &str) -> Vec<Chunk> {
    compile_policy(profile, body, case, Directives::default())
}

fn compile_policy(
    profile: LanguageProfile,
    body: Vec<Statement>,
    case: &str,
    extra: Directives,
) -> Vec<Chunk> {
    let mut directives = Directives {
        operator_dispatch: Some(OperatorDispatch::StaticBuiltin),
        type_resolution: Some(TypeResolution::Static),
        comparison_results_are_boolean: Some(true),
        logical_ops_bitwise_for_integers: Some(false),
        word_local_transport: Some(vybe_ast::WordLocalTransport::Boxed),
        ..Default::default()
    };
    directives.overlay(&extra);
    let module = Module {
        name: "<static-word-operand-requests>".into(),
        language: Lang::Unknown,
        body,
        imports: vec![],
        canon: Default::default(),
        directives,
    };
    let chunks = vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap();
    validate(&chunks, case);
    chunks
}

fn compile_result(profile: LanguageProfile, value: Expression) -> Vec<Chunk> {
    compile(profile, vec![ret(value)])
}

fn dump(chunks: &[Chunk]) -> String {
    chunks
        .iter()
        .map(vybe_runtime::debug::disassemble)
        .collect::<Vec<_>>()
        .join("\n")
}

fn narrowing_count(chunks: &[Chunk]) -> usize {
    dump(chunks).matches("i64.trunc_f64_s").count()
}

fn validate(chunks: &[Chunk], case: &str) {
    assert!(!dump(chunks).contains("dyneq"));
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: false,
        },
    );
    if let Err(error) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
    {
        let test = std::thread::current()
            .name()
            .unwrap_or("unnamed")
            .to_owned();
        let stem = format!("static-word-{test}-{case}-{}", std::process::id());
        let path = std::env::temp_dir().join(format!("{stem}-invalid.wasm"));
        std::fs::write(&path, &bytes).expect("save failing Wasm");
        std::fs::write(path.with_extension("chunks.txt"), dump(chunks))
            .expect("save failing Chunk context");
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let Ok(wasmparser::Payload::CodeSectionEntry(body)) = payload
                && body.range().contains(&error.offset())
            {
                let mut operators = body.get_operators_reader().unwrap();
                let mut previous = std::collections::VecDeque::new();
                while !operators.eof() {
                    let offset = operators.original_position();
                    let operation = operators.read().unwrap();
                    previous.push_back(format!("{offset}: {operation:?}"));
                    if previous.len() > 20 {
                        previous.pop_front();
                    }
                    if offset >= error.offset() {
                        eprintln!(
                            "failed body {:?}:\n{}",
                            body.range(),
                            previous.into_iter().collect::<Vec<_>>().join("\n")
                        );
                        break;
                    }
                }
            }
        }
        panic!(
            "test={test} case={case}: {error}; failing metadata-free Wasm: {}",
            path.display()
        );
    }
    let mut abstract_imports = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(imports) => {
                for import in imports.into_imports() {
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
}

type Trace = Arc<Mutex<Vec<String>>>;

fn vm(left: f64, right: f64) -> (VM, Trace) {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.memory.grow(1);
    vm.memory.with_buffer_mut(|bytes| bytes[0] = 9);
    let trace = Arc::new(Mutex::new(Vec::new()));
    for (name, value) in [("left", left), ("right", right), ("replace", 12.75)] {
        let trace = trace.clone();
        vm.register_host_fn(
            "test:static-word",
            name,
            Box::new(move |_, args| {
                assert!(args.is_empty());
                trace.lock().unwrap().push(name.into());
                Value::F64(value)
            }),
        );
    }
    (vm, trace)
}

fn legacy_probe(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    if !matches!(name, "static_word_raw.fetch" | "static_word_unknown.fetch") {
        return false;
    }
    assert_eq!(argc, 1);
    vybe_compiler::primitives::pointers::emit_linear_i32_load8_u(chunks, current, line);
    true
}

fn typed_probe(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
    request: EmitRequest,
) -> EmitOutcome {
    if !legacy_probe(name, chunks, current, argc, line) {
        return EmitOutcome::Unhandled;
    }
    EmitOutcome::Handled(
        if name == "static_word_raw.fetch" && request == EmitRequest::I32Word {
            ResultRepresentation::RawI32Word
        } else {
            ResultRepresentation::Unknown
        },
    )
}

fn parser(_: &str) -> Result<Module, String> {
    unreachable!("AST-only test provider")
}
fn provider_profile() -> &'static str {
    "[compiler]\n"
}

fn register_probes() {
    static REGISTER: Once = Once::new();
    REGISTER.call_once(|| {
        for namespace in ["static_word_raw", "static_word_unknown"] {
            registry::register_language(registry::LanguageDef {
                name: namespace,
                parse: parser,
                profile_source: provider_profile,
                emit_dispatch: Some(legacy_probe),
                normalize_class: None,
                register_tree: None,
                expand_source: None,
            });
            emitter_results::register(TypedEmitterDef {
                namespace,
                legacy_dispatch: legacy_probe,
                dispatch: typed_probe,
            })
            .unwrap();
        }
    });
}

#[test]
fn actual_producer_reports_not_integer_metadata_or_emitted_opcode_select_reuse() {
    register_probes();
    for (operation, expected_narrowings) in [
        ("pointers.linear_i32_load8_u", 0),
        ("static_word_raw.fetch", 0),
        ("static_word_unknown.fetch", 1),
    ] {
        let mut p = profile();
        p.builtins.get_mut("fetch").unwrap().emit = BuiltinEmit::Common(operation.into());
        for reverse in [false, true] {
            let read = call("fetch", vec![Expression::int(0)]);
            let (left, right) = if reverse {
                (Expression::int(1), read)
            } else {
                (read, Expression::int(1))
            };
            let chunks = compile_result(p.clone(), binary(BinOp::Add, left, right));
            assert_eq!(
                narrowing_count(&chunks),
                expected_narrowings,
                "{operation}: {}",
                dump(&chunks)
            );
            assert_eq!(vm(0.0, 0.0).0.run(chunks).unwrap().as_i32(), 10);
        }
    }
}

#[test]
fn unknown_host_fractional_results_keep_truncation_and_word_wrap() {
    for number in [-1.75, 12.75, 4_294_967_295.75] {
        for reverse in [false, true] {
            let (left, right) = if reverse {
                (Expression::int(1), call("left", vec![]))
            } else {
                (call("left", vec![]), Expression::int(1))
            };
            let chunks = compile_result(profile(), binary(BinOp::Add, left, right));
            assert_eq!(narrowing_count(&chunks), 1);
            let (mut runtime, trace) = vm(number, 0.0);
            assert_eq!(
                runtime.run(chunks).unwrap().as_i32(),
                (number as i64 as i32).wrapping_add(1)
            );
            assert_eq!(*trace.lock().unwrap(), ["left"]);
        }
    }
}

#[test]
fn nonfinite_fallback_traps_after_both_operands_evaluate() {
    for number in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY, 1e30] {
        for reverse in [false, true] {
            let chunks = compile_result(
                profile(),
                binary(BinOp::Add, call("left", vec![]), call("right", vec![])),
            );
            assert_eq!(narrowing_count(&chunks), 2);
            let (mut runtime, trace) = if reverse {
                vm(1.0, number)
            } else {
                vm(number, 1.0)
            };
            let error = runtime
                .run(chunks)
                .expect_err("nonfinite/out-of-range normalization must trap");
            let expected_trap = if number.is_nan() {
                "invalid conversion"
            } else {
                "integer overflow"
            };
            assert!(error.to_string().contains(expected_trap), "{error}");
            assert_eq!(*trace.lock().unwrap(), ["left", "right"]);
        }
    }
}

#[test]
fn nested_comparison_argument_remains_bool_and_outer_unknown_result_is_normalized() {
    for matches in [false, true] {
        let predicate = binary(BinOp::Eq, call("left", vec![]), Expression::int(0));
        let chunks = compile_result(
            profile(),
            binary(
                BinOp::Add,
                call("fetch", vec![Expression::int(0)]),
                call("identity", vec![predicate]),
            ),
        );
        assert_eq!(narrowing_count(&chunks), 2);
        let (mut runtime, trace) = vm(if matches { 0.0 } else { 1.0 }, 0.0);
        let observed = trace.clone();
        runtime.register_host_fn(
            "test:static-word",
            "identity",
            Box::new(move |_, args| {
                assert_eq!(
                    args,
                    &[Value::Bool(matches)],
                    "nested condition request leaked into value argument"
                );
                observed.lock().unwrap().push("identity".into());
                Value::F64(if matches { 1.0 } else { 0.0 })
            }),
        );
        assert_eq!(
            runtime.run(chunks).unwrap().as_i32(),
            9 + i32::from(matches)
        );
        assert_eq!(*trace.lock().unwrap(), ["left", "identity"]);
    }
}

#[test]
fn nested_raw_argument_cannot_certify_an_unknown_outer_call_result() {
    let chunks = compile_result(
        profile(),
        binary(
            BinOp::Add,
            call("identity", vec![call("fetch", vec![Expression::int(0)])]),
            Expression::int(1),
        ),
    );
    assert_eq!(narrowing_count(&chunks), 1);
    let (mut runtime, trace) = vm(0.0, 0.0);
    let observed = trace.clone();
    runtime.register_host_fn(
        "test:static-word",
        "identity",
        Box::new(move |_, args| {
            assert_eq!(args.len(), 1);
            assert_eq!(args[0].as_i32(), 9);
            observed.lock().unwrap().push("identity".into());
            Value::F64(f64::NAN)
        }),
    );
    assert!(runtime.run(chunks).is_err());
    assert_eq!(*trace.lock().unwrap(), ["identity"]);
}

#[test]
fn side_effecting_operands_run_left_to_right_once_in_values_and_conditions() {
    for (op, expected) in [
        (BinOp::Add, Value::I32(5)),
        (BinOp::Eq, Value::Bool(false)),
        (BinOp::Lt, Value::Bool(true)),
    ] {
        let chunks = compile_result(
            profile(),
            binary(op, call("left", vec![]), call("right", vec![])),
        );
        let (mut runtime, trace) = vm(2.0, 3.0);
        let result = runtime.run(chunks).unwrap();
        if matches!(expected, Value::Bool(_)) {
            assert_eq!(result, expected);
        } else {
            assert_eq!(result.as_i32(), expected.as_i32());
        }
        assert_eq!(*trace.lock().unwrap(), ["left", "right"]);
    }
    let chunks = compile(
        profile(),
        vec![Statement::new(StmtKind::If {
            cond: binary(BinOp::Lt, call("left", vec![]), call("right", vec![])),
            then_body: vec![ret(Expression::int(10))],
            elifs: vec![],
            else_body: Some(vec![ret(Expression::int(20))]),
        })],
    );
    let (mut runtime, trace) = vm(2.0, 3.0);
    assert_eq!(runtime.run(chunks).unwrap().as_i32(), 10);
    assert_eq!(*trace.lock().unwrap(), ["left", "right"]);
}

#[test]
fn nested_short_circuit_comparisons_keep_bool_values_and_skip_rhs_effects() {
    for left in [0.0, 1.0] {
        for right in [0.0, 1.0] {
            for op in [BinOp::And, BinOp::Or] {
                let chunks = compile_result(
                    profile(),
                    binary(
                        op,
                        binary(BinOp::Eq, call("left", vec![]), Expression::int(1)),
                        binary(BinOp::Eq, call("right", vec![]), Expression::int(1)),
                    ),
                );
                let expected = if op == BinOp::And {
                    left == 1.0 && right == 1.0
                } else {
                    left == 1.0 || right == 1.0
                };
                let skip = if op == BinOp::And {
                    left == 0.0
                } else {
                    left == 1.0
                };
                let (mut runtime, trace) = vm(left, right);
                assert_eq!(runtime.run(chunks).unwrap(), Value::Bool(expected));
                assert_eq!(
                    *trace.lock().unwrap(),
                    if skip {
                        vec!["left"]
                    } else {
                        vec!["left", "right"]
                    }
                );
            }
        }
    }
}

#[test]
fn side_effecting_checked_operand_and_rebound_callee_cannot_reuse_builtin_report() {
    use vybe_ast::{BindingPattern, VarDeclKind, VarDeclarator};
    let lambda = |value| {
        Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Expr(Box::new(value)),
            is_async: false,
            captures: vec![],
        })
    };
    for (type_name, expected) in [("int", 11.0), ("double", 4_294_967_307.75)] {
        let operand = Expression::new(ExprKind::Cast {
            type_name: type_name.into(),
            expr: Box::new(Expression::new(ExprKind::Sequence(vec![
                Expression::new(ExprKind::Assign {
                    target: Box::new(Expression::ident("fetch")),
                    value: Box::new(lambda(call("replace", vec![]))),
                }),
                call("fetch", vec![]),
            ]))),
        });
        let chunks = compile_case(
            profile(),
            vec![
                Statement::new(StmtKind::VarDecl {
                    kind: VarDeclKind::Let,
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident("fetch".into()),
                        type_hint: None,
                        init: Some(lambda(Expression::float(99.5))),
                        array_bounds: None,
                        with_events: false,
                    }],
                }),
                ret(binary(BinOp::Add, Expression::int(0xffff_ffff), operand)),
            ],
            type_name,
        );
        assert!(
            !dump(&chunks).contains("i32.load8_u"),
            "shadowed builtin must not emit its old producer"
        );
        if type_name == "int" {
            assert!(narrowing_count(&chunks) > 0);
        } else {
            assert!(
                dump(&chunks).contains("f64.add"),
                "mutating double operand must retain float dispatch"
            );
        }
        let (mut runtime, trace) = vm(0.0, 0.0);
        assert_eq!(runtime.run(chunks).unwrap().as_f64(), expected);
        assert_eq!(*trace.lock().unwrap(), ["replace"]);
    }
}

#[test]
fn rebound_callee_value_only_control_preserves_source_callable_without_word_request() {
    use vybe_ast::{BindingPattern, VarDeclKind, VarDeclarator};
    let lambda = |value| {
        Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Expr(Box::new(value)),
            is_async: false,
            captures: vec![],
        })
    };
    // No binary operation, raw literal, integer cast, or I32Word consumer.
    // Retain this control independently of the original int/double regression.
    let chunks = compile_case(
        profile(),
        vec![
            Statement::new(StmtKind::VarDecl {
                kind: VarDeclKind::Let,
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident("fetch".into()),
                    type_hint: None,
                    init: Some(lambda(Expression::float(99.5))),
                    array_bounds: None,
                    with_events: false,
                }],
            }),
            ret(Expression::new(ExprKind::Sequence(vec![
                Expression::new(ExprKind::Assign {
                    target: Box::new(Expression::ident("fetch")),
                    value: Box::new(lambda(call("replace", vec![]))),
                }),
                call("fetch", vec![]),
            ]))),
        ],
        "value-only-no-word-request",
    );
    assert!(!dump(&chunks).contains("i32.load8_u"));
    let (mut runtime, trace) = vm(0.0, 0.0);
    assert_eq!(runtime.run(chunks).unwrap().as_f64(), 12.75);
    assert_eq!(*trace.lock().unwrap(), ["replace"]);
}

fn drift_declaration(
    params: Vec<vybe_ast::Param>,
    return_type: &str,
    result: Expression,
) -> Statement {
    Statement::new(StmtKind::FunctionDecl {
        name: "f".into(),
        params,
        return_type: Some(return_type.into()),
        body: vec![ret(result)],
        modifiers: Default::default(),
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    })
}

#[test]
fn nested_declaration_target_drift_preserves_high_literal_until_final_dispatch() {
    let callback = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(vec![
            drift_declaration(vec![], "double", Expression::float(0.5)),
            ret(Expression::int(0)),
        ]),
        is_async: false,
        captures: vec![],
    });
    assert_drift_result(callback);
}

fn assert_drift_result(callback: Expression) {
    use vybe_ast::{Param, PassBy};
    let chunks = compile(
        profile(),
        vec![
            drift_declaration(
                vec![Param {
                    name: "callback".into(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }],
                "int",
                Expression::int(1),
            ),
            // Compiling the argument's nested f publishes a double return hint in
            // the shared inference map. Preserve the existing late target choice;
            // pre-emitting this literal as i32(-1) would turn the result into zero.
            ret(binary(
                BinOp::Add,
                Expression::int(0xffff_ffff),
                call("f", vec![callback]),
            )),
        ],
    );
    let root = vybe_runtime::debug::disassemble(&chunks[0]);
    assert!(
        root.contains("f64.add"),
        "late double target must remain selected: {root}"
    );
    assert!(
        root.contains("f64.const 4294967295"),
        "high literal was narrowed before final target resolution: {root}"
    );
    assert_eq!(
        vm(0.0, 0.0).0.run(chunks).unwrap().as_f64(),
        4_294_967_296.0
    );
}

fn skipped(value: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::int(0)),
        then: Box::new(value),
        else_: Box::new(Expression::int(0)),
    })
}

fn default_parameter() -> vybe_ast::Param {
    vybe_ast::Param {
        name: "item".into(),
        type_hint: None,
        default: Some(Expression::new(ExprKind::FunctionExpr(Box::new(
            drift_declaration(vec![], "double", Expression::float(0.5)),
        )))),
        pass_by: vybe_ast::PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: true,
        is_nullable: false,
    }
}

#[test]
fn class_expression_method_publication_cannot_narrow_parent_operand_early() {
    assert_drift_result(skipped(Expression::new(ExprKind::ClassExpr {
        name: Some("DeferredMetadata".into()),
        parent: None,
        interfaces: vec![],
        members: vec![vybe_ast::ClassMember::Method(Box::new(drift_declaration(
            vec![],
            "double",
            Expression::float(0.5),
        )))],
    })));
}

#[test]
fn array_map_synthetic_lambda_defaults_cannot_hide_target_drift() {
    assert_drift_result(skipped(Expression::new(ExprKind::ArrayMap {
        array: Box::new(Expression::new(ExprKind::Array(vec![]))),
        params: vec![default_parameter()],
        body: Box::new(Expression::int(0)),
    })));
}

#[test]
fn callable_adapter_synthetic_lambda_defaults_cannot_hide_target_drift() {
    assert_drift_result(skipped(Expression::new(ExprKind::CallableRef {
        target: Box::new(Expression::ident("f")),
        receiver: None,
        binding: vybe_ast::CallableBinding::Static,
        adapter: Some(vybe_ast::CallableAdapter::Expr {
            params: vec![default_parameter()],
            body: Box::new(Expression::int(0)),
        }),
    })));
}

#[test]
fn object_method_body_nested_declaration_cannot_hide_target_drift() {
    let method = Statement::new(StmtKind::FunctionDecl {
        name: "method".into(),
        params: vec![],
        return_type: None,
        body: vec![
            drift_declaration(vec![], "double", Expression::float(0.5)),
            ret(Expression::int(0)),
        ],
        modifiers: Default::default(),
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    });
    assert_drift_result(skipped(Expression::new(ExprKind::Object(vec![
        vybe_ast::ObjectProperty::Method {
            key: "method".into(),
            value: Box::new(method),
        },
    ]))));
}

fn shadow_binding(name: &str, kind: vybe_ast::VarDeclKind, answer: i64) -> Statement {
    Statement::new(StmtKind::VarDecl {
        kind,
        declarations: vec![vybe_ast::VarDeclarator {
            pattern: vybe_ast::BindingPattern::Ident(name.into()),
            type_hint: None,
            init: Some(Expression::new(ExprKind::Lambda {
                params: vec![vybe_ast::Param {
                    name: "address".into(),
                    type_hint: None,
                    default: None,
                    pass_by: vybe_ast::PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }],
                body: LambdaBody::Expr(Box::new(Expression::int(answer))),
                is_async: false,
                captures: vec![],
            })),
            array_bounds: None,
            with_events: false,
        }],
    })
}

fn probe_function(name: &str, body: Vec<Statement>) -> Statement {
    Statement::new(StmtKind::FunctionDecl {
        name: name.into(),
        params: vec![],
        return_type: None,
        body,
        modifiers: Default::default(),
        handles: vec![],
        is_async: false,
        is_generator: false,
        is_sub: false,
    })
}

fn scope_statement(kind: vybe_ast::ScopeDeclKind, names: &[&str]) -> Statement {
    Statement::new(StmtKind::ScopeDecl {
        kind,
        names: names.iter().map(|name| (*name).into()).collect(),
    })
}

#[test]
fn local_and_static_local_callables_shadow_profile_builtin() {
    for kind in [vybe_ast::VarDeclKind::Let, vybe_ast::VarDeclKind::Static] {
        let chunks = compile_case(
            profile(),
            vec![
                probe_function(
                    "probe",
                    vec![
                        shadow_binding("fetch", kind.clone(), 42),
                        ret(call("fetch", vec![Expression::int(0)])),
                    ],
                ),
                ret(call("probe", vec![])),
            ],
            &format!("local-{kind:?}"),
        );
        assert!(!dump(&chunks).contains("i32.load8_u"));
        assert_eq!(vm(0.0, 0.0).0.run(chunks).unwrap().as_i32(), 42);
    }
}

#[test]
fn closed_scope_ignores_module_binding_unless_reopened_or_locally_shadowed() {
    for (reopen, local, expected) in [(false, false, 9), (true, false, 81), (false, true, 42)] {
        let mut body = vec![scope_statement(vybe_ast::ScopeDeclKind::Closed, &[])];
        if reopen {
            body.push(scope_statement(vybe_ast::ScopeDeclKind::Global, &["fetch"]));
        }
        if local {
            body.push(shadow_binding("fetch", vybe_ast::VarDeclKind::Let, 42));
        }
        body.push(ret(call("fetch", vec![Expression::int(0)])));
        let chunks = compile_case(
            profile(),
            vec![
                shadow_binding("fetch", vybe_ast::VarDeclKind::Let, 81),
                probe_function("probe", body),
                ret(call("probe", vec![])),
            ],
            &format!("closed-reopen-{reopen}-local-{local}"),
        );
        assert_eq!(dump(&chunks).contains("i32.load8_u"), !reopen && !local);
        assert_eq!(vm(0.0, 0.0).0.run(chunks).unwrap().as_i32(), expected);
    }
}

#[test]
fn closed_scope_keeps_its_own_static_callable_visible() {
    let chunks = compile_case(
        profile(),
        vec![
            probe_function(
                "probe",
                vec![
                    scope_statement(vybe_ast::ScopeDeclKind::Closed, &[]),
                    shadow_binding("fetch", vybe_ast::VarDeclKind::Static, 42),
                    ret(call("fetch", vec![Expression::int(0)])),
                ],
            ),
            ret(call("probe", vec![])),
        ],
        "closed-own-static",
    );
    assert!(
        !dump(&chunks).contains("i32.load8_u"),
        "a closed scope must still dispatch its own static callable"
    );
    assert_eq!(vm(0.0, 0.0).0.run(chunks).unwrap().as_i32(), 42);
}

#[test]
fn closed_inner_scope_does_not_inherit_outer_static_builtin_shadow() {
    let chunks = compile_case(
        profile(),
        vec![
            probe_function(
                "outer",
                vec![
                    shadow_binding("fetch", vybe_ast::VarDeclKind::Static, 42),
                    probe_function(
                        "inner",
                        vec![
                            scope_statement(vybe_ast::ScopeDeclKind::Closed, &[]),
                            ret(call("fetch", vec![Expression::int(0)])),
                        ],
                    ),
                    ret(call("inner", vec![])),
                ],
            ),
            ret(call("outer", vec![])),
        ],
        "closed-inner-outer-static",
    );
    assert!(
        dump(&chunks).contains("i32.load8_u"),
        "invisible outer static must not suppress builtin"
    );
    assert_eq!(vm(0.0, 0.0).0.run(chunks).unwrap().as_i32(), 9);
}

#[test]
fn module_builtin_shadow_uses_declared_variable_case_policy() {
    for (policy, expected) in [
        (vybe_ast::CaseMatch::Exact, 9),
        (vybe_ast::CaseMatch::Folded, 81),
    ] {
        let chunks = compile_policy(
            profile(),
            vec![
                shadow_binding("FeTcH", vybe_ast::VarDeclKind::Let, 81),
                ret(call("fetch", vec![Expression::int(0)])),
            ],
            &format!("module-case-{policy:?}"),
            Directives {
                variable_case: Some(policy),
                case_alphabet: Some(vybe_ast::CaseAlphabet::Ascii),
                ..Default::default()
            },
        );
        assert_eq!(
            dump(&chunks).contains("i32.load8_u"),
            policy == vybe_ast::CaseMatch::Exact
        );
        assert_eq!(vm(0.0, 0.0).0.run(chunks).unwrap().as_i32(), expected);
    }
}

#[test]
fn unrelated_source_binding_does_not_hide_nonshadowed_builtin() {
    let chunks = compile_case(
        profile(),
        vec![
            shadow_binding("different", vybe_ast::VarDeclKind::Let, 81),
            ret(call("fetch", vec![Expression::int(0)])),
        ],
        "nonshadowed-builtin",
    );
    assert!(dump(&chunks).contains("i32.load8_u"));
    assert_eq!(vm(0.0, 0.0).0.run(chunks).unwrap().as_i32(), 9);
}

#[test]
fn separate_variable_namespace_keeps_bare_builtin_and_variable_callable_distinct() {
    static NAMESPACE: registry::VariableNamespace = registry::VariableNamespace {
        body: |name| name.strip_prefix('@').unwrap_or(name),
        spell: |name| format!("@{name}"),
        global_key: |_, body| Some(format!("variable:{body}")),
    };
    static REGISTER: Once = Once::new();
    REGISTER.call_once(|| {
        registry::register_hooks(
            "static_word_separate_namespace_control",
            registry::LanguageHooks {
                variable_namespace: Some(&NAMESPACE),
                ..Default::default()
            },
        )
    });
    let mut p = profile();
    p.name = "static_word_separate_namespace_control".into();
    // Deliberately register the variable spelling as a builtin too. Its explicit
    // variable-namespace binding must win, while bare fetch remains intrinsic.
    let builtin = p.builtins.get("fetch").unwrap().clone();
    p.builtins.insert("@fetch".into(), builtin);
    let chunks = compile_case(
        p,
        vec![
            shadow_binding("@fetch", vybe_ast::VarDeclKind::Let, 81),
            Statement::new(StmtKind::Expr(call("fetch", vec![Expression::int(0)]))),
            ret(call("@fetch", vec![Expression::int(0)])),
        ],
        "separate-variable-namespace",
    );
    assert_eq!(dump(&chunks).matches("i32.load8_u").count(), 1);
    assert_eq!(vm(0.0, 0.0).0.run(chunks).unwrap().as_i32(), 81);
}
