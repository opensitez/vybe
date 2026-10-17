#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::*;
use vybe_runtime::{Chunk, Value, VM};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn compiler(mode: u8) -> vybe_compiler::Compiler {
    let mut profile=vybe_compiler::profile::parse_profile("[compiler]\nhas_generators=true\necma_promise_methods=true\necma_iterator_result_shape=true\nhas_undefined_value=true\ncase_sensitive=true\nmaterialize_bool_results=true\n").unwrap();
    profile.buffered_iterator_methods = mode >= 5;
    let mut c = vybe_compiler::Compiler::with_profile(profile);
    c.benchmark_generator_control(mode > 0, mode > 1);
    c.benchmark_condition_truthiness_reuse(mode != 3);
    c.benchmark_integer_truthiness_reuse(!matches!(mode, 4 | 6));
    c
}
fn stmt(k: StmtKind) -> Statement {
    Statement::new(k)
}
fn expr(k: ExprKind) -> Expression {
    Expression::new(k)
}
fn lit(n: i64) -> Expression {
    Expression::int(n)
}
fn text(t: &str) -> Expression {
    expr(ExprKind::Lit(Literal::Str(t.into())))
}
fn id(t: &str) -> Expression {
    Expression::ident(t)
}
fn ret(e: Expression) -> Statement {
    stmt(StmtKind::Return(Some(e)))
}
fn yield_(e: Expression) -> Statement {
    stmt(StmtKind::Expr(expr(ExprKind::Yield(Some(Box::new(e))))))
}
fn set_trace(n: i64) -> Statement {
    stmt(StmtKind::Expr(expr(ExprKind::Assign {
        target: Box::new(id("trace")),
        value: Box::new(lit(n)),
    })))
}
fn call(method: &str, args: Vec<Expression>) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(expr(ExprKind::Member {
            object: Box::new(id("it")),
            field: method.into(),
            null_safe: false,
        })),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}
fn function(name: &str, params: &[&str], body: Vec<Statement>, generator: bool) -> Statement {
    stmt(StmtKind::FunctionDecl {
        name: name.into(),
        params: params
            .iter()
            .map(|name| Param {
                name: (*name).into(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect(),
        body,
        return_type: None,
        modifiers: Modifiers::default(),
        handles: vec![],
        is_async: false,
        is_generator: generator,
        is_sub: false,
    })
}
fn program(mode: u8, finally_yields: bool, yields: usize) -> Vec<Chunk> {
    let mut tail = vec![set_trace(7)];
    if finally_yields {
        tail.push(yield_(lit(9)));
        tail.push(set_trace(17));
    }
    let mut producer: Vec<_> = (1..=yields).map(|n| yield_(lit(n as i64))).collect();
    producer.push(ret(lit(300)));
    let body = vec![
        stmt(StmtKind::VarDecl {
            kind: VarDeclKind::Var,
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("trace".into()),
                type_hint: None,
                init: Some(lit(0)),
                array_bounds: None,
                with_events: false,
            }],
        }),
        function(
            "make",
            &[],
            vec![stmt(StmtKind::Try {
                body: producer,
                catches: vec![],
                else_body: None,
                finally: Some(tail),
            })],
            true,
        ),
        function("step", &["it"], vec![ret(call("next", vec![]))], false),
        function(
            "send",
            &["it", "arg"],
            vec![ret(call("next", vec![id("arg")]))],
            false,
        ),
        function(
            "stop",
            &["it", "arg"],
            vec![ret(call("return", vec![id("arg")]))],
            false,
        ),
        function(
            "fail",
            &["it"],
            vec![
                stmt(StmtKind::Try {
                    body: vec![stmt(StmtKind::Expr(call("throw", vec![text("boom")])))],
                    catches: vec![CatchClause {
                        types: vec![],
                        var_name: Some("e".into()),
                        stack_var: None,
                        body: vec![ret(text("caught"))],
                        when_clause: None,
                    }],
                    else_body: None,
                    finally: None,
                }),
                ret(text("not-caught")),
            ],
            false,
        ),
        function(
            "send_effect",
            &["it"],
            vec![ret(call(
                "next",
                vec![expr(ExprKind::Assign {
                    target: Box::new(id("trace")),
                    value: Box::new(lit(23)),
                })],
            ))],
            false,
        ),
        function(
            "stop_effect",
            &["it"],
            vec![ret(call(
                "return",
                vec![expr(ExprKind::Assign {
                    target: Box::new(id("trace")),
                    value: Box::new(lit(29)),
                })],
            ))],
            false,
        ),
        function(
            "send_close",
            &["it"],
            vec![ret(call("next", vec![call("return", vec![lit(19)])]))],
            false,
        ),
        function(
            "consume",
            &[],
            vec![
                stmt(StmtKind::VarDecl {
                    kind: VarDeclKind::Var,
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident("sum".into()),
                        type_hint: None,
                        init: Some(lit(0)),
                        array_bounds: None,
                        with_events: false,
                    }],
                }),
                stmt(StmtKind::ForIn {
                    var: "item".into(),
                    key: None,
                    iter: expr(ExprKind::Call {
                        callee: Box::new(id("make")),
                        args: vec![],
                        optional: false,
                    }),
                    body: vec![stmt(StmtKind::Expr(expr(ExprKind::Assign {
                        target: Box::new(id("sum")),
                        value: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(id("sum")),
                            right: Box::new(id("item")),
                        })),
                    })))],
                    of: true,
                    else_body: None,
                    is_async: false,
                }),
                ret(id("sum")),
            ],
            false,
        ),
        function("read_trace", &[], vec![ret(id("trace"))], false),
        ret(expr(ExprKind::Object(
            [
                "make",
                "step",
                "send",
                "stop",
                "fail",
                "send_effect",
                "stop_effect",
                "send_close",
                "consume",
                "read_trace",
            ]
            .into_iter()
            .map(|name| ObjectProperty::KeyValue {
                key: text(name),
                value: id(name),
            })
            .collect(),
        ))),
    ];
    let module = Module {
        name: "generator-lifecycle".into(),
        language: Lang::Unknown,
        imports: vec![],
        canon: Default::default(),
        directives: Directives::default(),
        body,
    };
    compiler(mode).compile(&module).unwrap()
}
fn property(v: &Value, k: &str) -> Value {
    let Value::Object(o) = v else {
        panic!("object: {v:?}")
    };
    let o = o.lock().unwrap();
    o.properties
        .get(k)
        .unwrap_or_else(|| panic!("missing {k}: {o:?}"))
        .clone()
}
fn setup(mode: u8, fy: bool, yields: usize) -> (VM, Value) {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let api = vm.run(program(mode, fy, yields)).unwrap();
    (vm, api)
}
fn invoke(vm: &mut VM, api: &Value, name: &str, args: &[Value]) -> Value {
    vm.invoke(&property(api, name), args).unwrap()
}
fn completion(v: &Value, n: f64, done: bool) {
    assert_eq!(property(v, "value").as_f64(), n);
    assert_eq!(property(v, "done"), Value::Bool(done));
}
#[test]
fn real_compiler_generator_return_throw_finally_and_resume() {
    let mut checks = 0;
    for mode in [0, 1, 2, 3, 4] {
        // A return before the body starts must not enter finally.
        let (mut vm, api) = setup(mode, false, 2);
        let g = invoke(&mut vm, &api, "make", &[]);
        let r = invoke(&mut vm, &api, "stop", &[g.clone(), Value::I32(17)]);
        completion(&r, 17., true);
        assert_eq!(invoke(&mut vm, &api, "read_trace", &[]).as_f64(), 0.);
        checks += 2;
        // Return through a suspended try must run finally and preserve payload.
        let (mut vm, api) = setup(mode, false, 2);
        let g = invoke(&mut vm, &api, "make", &[]);
        completion(&invoke(&mut vm, &api, "step", &[g.clone()]), 1., false);
        completion(
            &invoke(&mut vm, &api, "stop", &[g.clone(), Value::I32(17)]),
            17.,
            true,
        );
        assert_eq!(invoke(&mut vm, &api, "read_trace", &[]).as_f64(), 7.);
        checks += 3;
        // A yield in finally suspends pending return without losing its value.
        let (mut vm, api) = setup(mode, true, 2);
        let g = invoke(&mut vm, &api, "make", &[]);
        completion(&invoke(&mut vm, &api, "step", &[g.clone()]), 1., false);
        completion(
            &invoke(&mut vm, &api, "stop", &[g.clone(), Value::I32(17)]),
            9.,
            false,
        );
        assert_eq!(invoke(&mut vm, &api, "read_trace", &[]).as_f64(), 7.);
        completion(&invoke(&mut vm, &api, "step", &[g.clone()]), 17., true);
        assert_eq!(invoke(&mut vm, &api, "read_trace", &[]).as_f64(), 17.);
        checks += 5;
        // Same-Source caught throw avoids retaining an unhandled host exception.
        let (mut vm, api) = setup(mode, false, 2);
        let g = invoke(&mut vm, &api, "make", &[]);
        completion(&invoke(&mut vm, &api, "step", &[g.clone()]), 1., false);
        assert_eq!(
            invoke(&mut vm, &api, "fail", &[g.clone()]),
            Value::String("caught".into())
        );
        assert_eq!(invoke(&mut vm, &api, "read_trace", &[]).as_f64(), 7.);
        assert_eq!(
            property(&invoke(&mut vm, &api, "step", &[g]), "done"),
            Value::Bool(true)
        );
        checks += 4;
        // Ordinary next/send paths execute all actual compiled resume points.
        let (mut vm, api) = setup(mode, false, 4);
        let g = invoke(&mut vm, &api, "make", &[]);
        for n in 1..=4 {
            let r = if n % 2 == 0 {
                invoke(&mut vm, &api, "send", &[g.clone(), Value::Undefined])
            } else {
                invoke(&mut vm, &api, "step", &[g.clone()])
            };
            completion(&r, n as f64, false);
            checks += 1;
        }
        completion(&invoke(&mut vm, &api, "step", &[g.clone()]), 300., true);
        assert_eq!(invoke(&mut vm, &api, "read_trace", &[]).as_f64(), 7.);
        checks += 2;
        // Completed generators still evaluate arguments, once, before state checks.
        let r = invoke(&mut vm, &api, "send_effect", &[g.clone()]);
        assert_eq!(property(&r, "value"), Value::Undefined);
        assert_eq!(property(&r, "done"), Value::Bool(true));
        assert_eq!(invoke(&mut vm, &api, "read_trace", &[]).as_f64(), 23.);
        completion(
            &invoke(&mut vm, &api, "stop_effect", &[g.clone()]),
            29.,
            true,
        );
        assert_eq!(invoke(&mut vm, &api, "read_trace", &[]).as_f64(), 29.);
        checks += 5;
        // The argument can close the same generator before outer next resumes it.
        let g = invoke(&mut vm, &api, "make", &[]);
        let r = invoke(&mut vm, &api, "send_close", &[g]);
        assert_eq!(property(&r, "value"), Value::Undefined);
        assert_eq!(property(&r, "done"), Value::Bool(true));
        checks += 2;
        // Exercise the direct-generator for-in condition and finally cleanup.
        assert_eq!(invoke(&mut vm, &api, "consume", &[]).as_f64(), 10.);
        assert_eq!(invoke(&mut vm, &api, "read_trace", &[]).as_f64(), 7.);
        checks += 2;
    }
    eprintln!(
        "Verified {checks} actual Compiler generator lifecycle outcomes across original, predicate-only, plain-packet and redundant-truthiness modes"
    );
}
#[test]
fn buffered_generator_for_in_integer_flags() {
    for mode in [5, 6] {
        for (yields, finally_yields, sum, trace) in
            [(0, false, 0., 7.), (4, false, 10., 7.), (4, true, 19., 17.)]
        {
            let (mut vm, api) = setup(mode, finally_yields, yields);
            assert_eq!(invoke(&mut vm, &api, "consume", &[]).as_f64(), sum);
            assert_eq!(invoke(&mut vm, &api, "read_trace", &[]).as_f64(), trace);
        }
    }
    eprintln!("Verified 6 buffered generator iteration outcomes including empty/completion/yielding-finally");
}

#[test]
fn packet_schema_identity_values_and_size() {
    let mut rows = vec![];
    for op in ["return", "throw"] {
        let mut sizes = vec![];
        for mode in [0, 2] {
            let cs = compiler(mode).benchmark_generator_packet_chunks(op, 1);
            let mut vm = VM::new();
            vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
            let f = vm.run(cs.clone()).unwrap();
            let payloads = [
                Value::I32(17),
                Value::Null,
                Value::Undefined,
                Value::Bool(false),
                Value::String("é😀".into()),
            ];
            for v in payloads {
                let a = vm.invoke(&f, &[v.clone()]).unwrap();
                let b = vm.invoke(&f, &[v.clone()]).unwrap();
                assert_ne!(a, b);
                assert_eq!(property(&a, "__vybe_generator_control"), Value::Bool(true));
                assert_eq!(property(&a, "op"), Value::String(op.into()));
                assert_eq!(property(&a, "value"), v);
                let Value::Object(handle) = a else {
                    unreachable!()
                };
                assert_eq!(
                    handle.lock().unwrap().properties.contains_key("__keys"),
                    mode == 0
                );
            }
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap();
            if std::env::var_os("VYBE_PRIMITIVE_COMPILER_GENERATOR_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("compiler-packet-{op}-{mode}"));
            }
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        rows.push(serde_json::json!({"op":op,"sizes":sizes}));
    }
    let mut baseline_validation = None;
    for mode in [0, 1, 2, 3, 4] {
        let cs = program(mode, false, 8);
        let bytes = vybe_platform_wasm::write_wasm_with_options(
            &cs,
            vybe_platform_wasm::WasmWriteOptions {
                include_name_section: false,
                include_vybe_metadata: false,
            },
        );
        // The full lifecycle currently exposes a writer type mismatch even in
        // baseline mode. Record it explicitly; do not certify native lifecycle.
        // Candidates must not introduce a different validation failure.
        let validation = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .map(|_| "valid".to_owned())
            .unwrap_or_else(|e| e.message().to_owned());
        if mode == 0 {
            baseline_validation = Some(validation.clone());
        } else {
            assert_eq!(Some(&validation), baseline_validation.as_ref());
        }
        rows.push(serde_json::json!({"lifecycle_mode":mode,"chunk_sizes":cs.iter().map(|c|serde_json::json!({"name":c.name,"body":c.code.len(),"locals":c.local_count})).collect::<Vec<_>>(),"native_validation":validation,"native_bytes_unvalidated":bytes.len(),"all_bodies":cs.iter().map(|c|c.code.len()).sum::<usize>(),"generator_body":cs.iter().filter(|c|c.name=="make").map(|c|c.code.len()).collect::<Vec<_>>() }));
    }
    let truthiness_only = std::env::var_os("VYBE_TRUTHINESS_ONLY").is_some();
    let integer_only = std::env::var_os("VYBE_INTEGER_ONLY").is_some();
    if integer_only
        || truthiness_only
        || std::env::var_os("VYBE_PRIMITIVE_COMPILER_GENERATOR_EXPORT").is_some()
    {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(if integer_only {
                "../../INTEGER_GENERATOR_GATES.json"
            } else if truthiness_only {
                "../../CONDITION_GENERATOR_GATES.json"
            } else {
                "../../COMPILER_GENERATOR_GATES.json"
            }),
            serde_json::to_vec_pretty(&rows).unwrap(),
        )
        .unwrap();
    }
    eprintln!("{rows:?}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    let truthiness_only = std::env::var_os("VYBE_TRUTHINESS_ONLY").is_some();
    let integer_only = std::env::var_os("VYBE_INTEGER_ONLY").is_some();
    let baseline = if integer_only {
        4
    } else if truthiness_only {
        3
    } else {
        0
    };
    for label in [
        "packet-construction",
        "fresh-close",
        "suspended-close",
        "normal-8-yields",
        "for-in-8-yields",
        "buffered-for-in-8-yields",
    ] {
        if label == "buffered-for-in-8-yields" && !integer_only {
            continue;
        }
        for candidate in if truthiness_only || integer_only {
            vec![2u8]
        } else {
            vec![1u8, 2]
        } {
            if (truthiness_only || integer_only) && label == "packet-construction" {
                continue;
            }
            if label == "packet-construction" && candidate == 1 {
                continue;
            }
            let mut workers = if label == "packet-construction" {
                [baseline, candidate].map(|mode| {
                    let mut vm = VM::new();
                    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
                    let f = vm
                        .run(compiler(mode).benchmark_generator_packet_chunks("return", 32))
                        .unwrap();
                    (vm, f)
                })
            } else {
                if label == "buffered-for-in-8-yields" {
                    [setup(6, false, 8), setup(5, false, 8)]
                } else {
                    [setup(baseline, false, 8), setup(candidate, false, 8)]
                }
            };
            let run = |w: &mut (VM, Value)| {
                let (vm, api) = w;
                if label == "packet-construction" {
                    std::hint::black_box(vm.invoke(api, &[Value::I32(17)]).unwrap());
                } else {
                    if label == "for-in-8-yields" || label == "buffered-for-in-8-yields" {
                        std::hint::black_box(invoke(vm, api, "consume", &[]));
                        return;
                    }
                    let g = invoke(vm, api, "make", &[]);
                    if label == "normal-8-yields" {
                        for _ in 0..9 {
                            std::hint::black_box(invoke(vm, api, "step", &[g.clone()]));
                        }
                    } else {
                        if label == "suspended-close" {
                            std::hint::black_box(invoke(vm, api, "step", &[g.clone()]));
                        }
                        std::hint::black_box(invoke(vm, api, "stop", &[g, Value::I32(17)]));
                    }
                }
            };
            for w in &mut workers {
                for _ in 0..8 {
                    run(w);
                }
            }
            let mut samples = [vec![], vec![]];
            for pair in 0..7 {
                for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let start = Instant::now();
                    let mut n = 0;
                    while start.elapsed() < Duration::from_millis(200) || n < 8 {
                        run(&mut workers[side]);
                        n += 1;
                    }
                    samples[side].push(start.elapsed().as_secs_f64() / n as f64);
                }
            }
            let mut ratios: Vec<_> = samples[0]
                .iter()
                .zip(&samples[1])
                .map(|(a, b)| a / b)
                .collect();
            ratios.sort_by(f64::total_cmp);
            rows.push(serde_json::json!({"case":label,"candidate_mode":candidate,"median_speedup":ratios[3],"seconds_per_call":samples}));
        }
    }
    let result = serde_json::json!({"pairs":7,"minimum_sample_ms":200,"minimum_calls":8,"warmups":8,"packet_constructions_per_call":32,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(if integer_only {
            "../../INTEGER_GENERATOR_MEASUREMENTS.json"
        } else if truthiness_only {
            "../../CONDITION_GENERATOR_MEASUREMENTS.json"
        } else {
            "../../COMPILER_GENERATOR_MEASUREMENTS.json"
        }),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
