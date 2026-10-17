#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::*;
use vybe_compiler::primitives::{self, ops};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn lit(n: i64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Int(n)))
}
fn param(name: &str, hint: Option<TypeHint>) -> Param {
    Param {
        name: name.into(),
        type_hint: hint,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}
fn call(name: &str) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: vec![Argument::positional(Expression::ident("x"))],
        optional: false,
    })
}
fn program(fast: bool, kind: usize, repeats: usize) -> Vec<Chunk> {
    ops::benchmark_with_proven_i32_conditions(fast, || {
        if kind == 6 {
            let mut root = Chunk::new("constructor-condition");
            primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
            root.emit_op(Op::RETURN, 0);
            let mut c = Chunk::new("entry");
            c.arity = 1;
            c.param_count = 1;
            c.local_count = 1;
            c.emit_op_u16(Op::LOCAL_GET, 0, 0);
            let mut cs = vec![root, c];
            primitives::collections::emit_spread_iterable_for_constructor(&mut cs, 1, 0);
            cs[1].emit_op(Op::RETURN, 0);
            primitives::bundle::finalize_with_runtime_helpers(&mut cs);
            primitives::globals::declare_free_globals(&mut cs);
            primitives::globals::normalize_global_table(&mut cs);
            return cs;
        }
        let profile = primitives_profile();
        let compiler = vybe_compiler::Compiler::with_profile(profile);
        let mut params = vec![param(
            "x",
            if kind == 5 {
                Some(TypeHint::converting("string"))
            } else {
                None
            },
        )];
        let expr = match kind {
            0 => call("is_float"),
            1 => call("isref"),
            2 => Expression::new(ExprKind::IsType {
                expr: Box::new(Expression::ident("x")),
                type_name: "integer".into(),
            }),
            _ => Expression::ident("x"),
        };
        let mut body = vec![];
        if kind == 3 {
            body.push(Statement::new(StmtKind::ReDim {
                preserve: true,
                array: "x".into(),
                bounds: vec![lit(3)],
            }));
        } else if kind == 4 || kind == 5 {
            params.push(param("y", None));
            body.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Index {
                    object: Box::new(Expression::ident("x")),
                    index: Box::new(Expression::new(ExprKind::Slice {
                        lower: Some(Box::new(lit(1))),
                        upper: Some(Box::new(lit(3))),
                        step: None,
                    })),
                    null_safe: false,
                })],
                value: Expression::ident("y"),
                by_ref: false,
            }));
        } else {
            for _ in 1..repeats {
                body.push(Statement::new(StmtKind::Expr(expr.clone())));
            }
        }
        body.push(Statement::new(StmtKind::Return(Some(expr))));
        compiler
            .compile(&Module {
                name: "proven-i32".into(),
                language: Lang::Unknown,
                imports: vec![],
                canon: Default::default(),
                directives: Default::default(),
                body: vec![
                    Statement::new(StmtKind::FunctionDecl {
                        name: "entry".into(),
                        params,
                        body,
                        return_type: None,
                        modifiers: Default::default(),
                        handles: vec![],
                        is_async: false,
                        is_generator: false,
                        is_sub: false,
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident("entry")))),
                ],
            })
            .unwrap()
    })
}
fn primitives_profile() -> vybe_compiler::profile::LanguageProfile {
    vybe_compiler::profile::parse_profile("[compiler]\ncase_sensitive=true\nslice_bounds_inclusive=true\n[builtins]\nis_float={emit=\"intrinsic:php_is_float\",min_args=1,max_args=1}\nisref={emit=\"intrinsic:isreference\",min_args=1,max_args=1}\n").unwrap()
}
fn worker(fast: bool, kind: usize) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, kind, 1)).unwrap();
    (vm, f)
}
fn array(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(
        vybe_runtime::value::Object::new_array(v),
    ))
}
fn snap(v: &Value) -> String {
    match v {
        Value::Object(o) => {
            let o = o.lock().unwrap();
            if let vybe_runtime::value::ObjectKind::Array(a) = &o.kind {
                format!("array:{:?}", a.iter().map(snap).collect::<Vec<_>>())
            } else {
                "object".into()
            }
        }
        _ => format!("{v:?}"),
    }
}
fn inputs(kind: usize) -> Vec<Vec<Value>> {
    let values = vec![
        Value::I32(0),
        Value::I32(7),
        Value::F64(0.5),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::F64(-0.0),
        Value::Null,
        Value::Undefined,
        Value::Bool(true),
        Value::Bool(false),
        Value::String("".into()),
        Value::String("7".into()),
        Value::bigint_i64(0),
        array(vec![]),
    ];
    match kind {
        0..=2 => values.into_iter().map(|x| vec![x]).collect(),
        3 => vec![
            vec![array(vec![Value::I32(7), Value::I32(2)])],
            vec![array(vec![Value::Bool(true)])],
            vec![array(vec![Value::String("s".into())])],
            vec![array(vec![])],
        ],
        4 => vec![
            vec![Value::String("abcd".into()), Value::String("Q".into())],
            vec![
                array((0..4).map(Value::I32).collect()),
                Value::String("Q".into()),
            ],
            vec![
                array((0..4).map(Value::I32).collect()),
                array(vec![Value::I32(8)]),
            ],
        ],
        5 => vec![
            vec![Value::String("abcd".into()), Value::String("Q".into())],
            vec![Value::String("abcd".into()), Value::String("LONG".into())],
        ],
        _ => vec![vec![Value::Null], vec![Value::Undefined]],
    }
}
#[test]
fn actual_lowerings_values_and_sizes() {
    let mut rows = vec![];
    let mut checks = 0;
    for kind in 0..7 {
        let mut ws = [worker(false, kind), worker(true, kind)];
        for case in 0..inputs(kind).len() {
            let mut out = vec![];
            for (vm, f) in &mut ws {
                let args = inputs(kind).remove(case);
                out.push(vm.invoke(f, &args).unwrap());
            }
            assert_eq!(snap(&out[0]), snap(&out[1]), "kind={kind} case={case}");
            if kind == 0 {
                let x = &inputs(kind)[case][0];
                let number = matches!(x, Value::I32(_) | Value::F64(_));
                let expected = number && (!x.as_f64().is_finite() || x.as_f64().fract() != 0.0);
                assert_eq!(
                    out[1],
                    if number {
                        Value::I32(i32::from(expected))
                    } else {
                        Value::Bool(false)
                    }
                );
            }
            if kind == 4 && case == 0 {
                assert_eq!(out[1], Value::String("aQ d".into()));
            }
            if kind == 6 {
                assert_eq!(snap(&out[1]), "array:[]");
            }
            checks += 1;
        }
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, kind, 1);
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
            let mut exported = cs.clone();
            let aliases: Vec<_> = exported
                .iter()
                .flat_map(|c| c.global_inits.iter())
                .filter_map(|g| {
                    if let vybe_runtime::chunk::ConstExpr::RefFunc(index) = g.init {
                        Some((g.name.clone(), index))
                    } else {
                        None
                    }
                })
                .collect();
            for (name, index) in aliases {
                let mut factory = Chunk::new(format!("binding-factory:{name}"));
                primitives::functions::emit_ref_func(&mut factory, index, 0, 0);
                factory.emit_op(Op::RETURN, 0);
                exported.push(factory);
            }
            native_export::export_chunks(&exported, &format!("proven-i32-{kind}-{fast}"));
            let body = cs.iter().find(|c| c.name == "entry").unwrap();
            sizes.push(serde_json::json!({"body":body.code.len(),"locals":body.local_count,"native":bytes.len(),"linked_body":cs.iter().map(|c|c.code.len()).sum::<usize>(),"linked_chunks":cs.len()}));
        }
        if kind == 2 {
            assert_eq!(
                sizes[0], sizes[1],
                "existing primitive type-query fast path control"
            );
        } else {
            assert!(
                sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64(),
                "kind={kind} {sizes:?}"
            );
        }
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../PROVEN_I32_GATES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
#[test]
fn paired_dispatch_provenance() {
    let mut rows = vec![];
    for kind in 0..=5 {
        let (mut vm, _) = worker(false, kind);
        let a = vm.run(program(false, kind, 1)).unwrap();
        let b = vm.run(program(true, kind, 1)).unwrap();
        let chunk_index = |v: &Value| {
            let Value::Object(o) = v else {
                panic!("expected function")
            };
            let o = o.lock().unwrap();
            let vybe_runtime::value::ObjectKind::Function(f) = &o.kind else {
                panic!("expected guest function")
            };
            f.chunk_index
        };
        let indices = [chunk_index(&a), chunk_index(&b)];
        assert_ne!(
            indices[0], indices[1],
            "old/new must call distinct loaded chunks"
        );
        let args = match kind {
            0 => vec![Value::F64(0.5)],
            1 => vec![Value::String("x".into())],
            2 => vec![Value::I32(7)],
            3 => vec![array(vec![Value::I32(7)])],
            _ => vec![Value::String("abcd".into()), Value::String("Q".into())],
        };
        let mut dispatches = vec![];
        for f in [&a, &b] {
            vm.invoke(f, &args).unwrap();
            vm.record_runtime_perf(true);
            vm.invoke(f, &args).unwrap();
            dispatches.push(vm.take_runtime_perf().unwrap().dispatches);
        }
        assert!(
            dispatches[1] <= dispatches[0],
            "kind={kind}: {dispatches:?}"
        );
        rows.push(serde_json::json!({"kind":kind,"chunk_indices":indices,"dispatches":dispatches}));
    }
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../PROVEN_I32_DISPATCHES.json"),
        serde_json::to_vec_pretty(&rows).unwrap(),
    )
    .unwrap();
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (kind, repeats, control) in [
        (0, 32, true),
        (0, 32, false),
        (1, 32, false),
        (2, 32, true),
        (3, 1, false),
        (4, 1, false),
        (5, 1, false),
    ] {
        let (mut vm, _) = worker(false, kind);
        let a = vm.run(program(false, kind, repeats)).unwrap();
        let b = if control {
            a.clone()
        } else {
            vm.run(program(true, kind, repeats)).unwrap()
        };
        let fs = [a, b];
        let args = match kind {
            0 => vec![Value::F64(0.5)],
            1 => vec![Value::String("x".into())],
            2 => vec![Value::I32(7)],
            3 => vec![array(vec![Value::I32(7)])],
            _ => vec![Value::String("abcd".into()), Value::String("Q".into())],
        };
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(vm.invoke(f, &args).unwrap());
            }
        }
        // Pair small equal-call batches so changing background load is shared
        // by both variants, instead of separating their samples by500ms.
        let calibration = Instant::now();
        for _ in 0..16 {
            std::hint::black_box(vm.invoke(&fs[0], &args).unwrap());
        }
        let batch_size =
            ((0.003 * 16.0 / calibration.elapsed().as_secs_f64()).ceil() as usize).clamp(1, 512);
        let mut samples = [vec![], vec![]];
        let mut calls_per_side = vec![];
        for pair in 0..7 {
            let mut seconds = [0.0f64, 0.0];
            let mut calls = 0;
            let mut turn = pair;
            while seconds[0] < 0.5 || seconds[1] < 0.5 || calls < 8 {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let start = Instant::now();
                    for _ in 0..batch_size {
                        std::hint::black_box(vm.invoke(&fs[side], &args).unwrap());
                    }
                    seconds[side] += start.elapsed().as_secs_f64();
                }
                calls += batch_size;
                turn += 1;
            }
            for side in 0..2 {
                samples[side].push(seconds[side] / calls as f64);
            }
            calls_per_side.push(calls);
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"kind":kind,"repeats":repeats,"control":control,"median_speedup":ratios[3],"batch_size":batch_size,"calls_per_side":calls_per_side,"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"same_vm":true,"interleaved_equal_call_batches":true,"batch_target_ms":3,"pairs":7,"minimum_sample_ms":500,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../PROVEN_I32_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
