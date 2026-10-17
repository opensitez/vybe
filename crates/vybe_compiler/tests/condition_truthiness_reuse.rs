#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::*;
use vybe_runtime::{Chunk, Value, VM};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn array(values: Vec<Value>) -> Value {
    Value::Object(std::sync::Arc::new(std::sync::Mutex::new(
        vybe_runtime::value::Object::new_array(values),
    )))
}
fn compiler(fast: bool) -> vybe_compiler::Compiler {
    let profile = vybe_compiler::profile::parse_profile(
        "[compiler]\nnegative_index_wraps=true\ncase_sensitive=true\n",
    )
    .unwrap();
    let mut c = vybe_compiler::Compiler::with_profile(profile);
    c.benchmark_condition_truthiness_reuse(fast);
    c
}
fn chunks(fast: bool, kind: usize, repeats: usize) -> Vec<Chunk> {
    if kind == 0 {
        return compiler(fast).benchmark_negated_condition_chunks(repeats);
    }
    let e = if kind == 3 {
        Expression::new(ExprKind::Index {
            object: Box::new(Expression::ident("a")),
            index: Box::new(Expression::ident("b")),
            null_safe: false,
        })
    } else {
        Expression::new(ExprKind::Binary {
            op: if kind == 1 { BinOp::Eqv } else { BinOp::Imp },
            left: Box::new(Expression::ident("a")),
            right: Box::new(Expression::ident("b")),
        })
    };
    let mut body: Vec<_> = (1..repeats)
        .map(|_| Statement::new(StmtKind::Expr(e.clone())))
        .collect();
    body.push(Statement::new(StmtKind::Return(Some(e))));
    compiler(fast)
        .compile(&Module {
            name: "condition-reuse".into(),
            language: Lang::Unknown,
            imports: vec![],
            canon: Default::default(),
            directives: Default::default(),
            body: vec![
                Statement::new(StmtKind::FunctionDecl {
                    name: "condition".into(),
                    params: ["a", "b"]
                        .into_iter()
                        .map(|n| Param {
                            name: n.into(),
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
                    modifiers: Default::default(),
                    handles: vec![],
                    is_async: false,
                    is_generator: false,
                    is_sub: false,
                }),
                Statement::new(StmtKind::Return(Some(Expression::ident("condition")))),
            ],
        })
        .unwrap()
}
fn worker(fast: bool, kind: usize, repeats: usize) -> (VM, Value) {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(chunks(fast, kind, repeats)).unwrap();
    (vm, f)
}
#[test]
fn actual_conditions_values_and_sizes() {
    let mut rows = vec![];
    let mut checks = 0;
    for kind in 0..4 {
        let mut ws = [worker(false, kind, 1), worker(true, kind, 1)];
        let cases: Vec<_> = if kind == 0 {
            vec![
                Value::Null,
                Value::Undefined,
                Value::Bool(false),
                Value::Bool(true),
                Value::I32(0),
                Value::I32(-1),
                Value::I32(2),
                Value::F64(-0.),
                Value::F64(f64::NAN),
                Value::F64(f64::INFINITY),
                Value::String("".into()),
                Value::String("0".into()),
                Value::String("é😀".into()),
                array(vec![]),
            ]
            .into_iter()
            .map(|v| vec![v])
            .collect()
        } else if kind == 3 {
            (-5..=5)
                .map(|i| {
                    vec![
                        array(vec![Value::I32(11), Value::I32(22), Value::I32(33)]),
                        Value::I32(i),
                    ]
                })
                .collect()
        } else {
            [-2, 0, 1, 2, 17]
                .into_iter()
                .flat_map(|a| {
                    [-2, 0, 1, 2, 17]
                        .into_iter()
                        .map(move |b| vec![Value::I32(a), Value::I32(b)])
                })
                .collect()
        };
        for args in cases {
            let f0 = ws[0].1.clone();
            let f1 = ws[1].1.clone();
            let a = ws[0].0.invoke(&f0, &args).unwrap();
            let b = ws[1].0.invoke(&f1, &args).unwrap();
            assert_eq!(a, b, "kind {kind}");
            if kind == 1 {
                assert_eq!(a.as_f64(), f64::from(args[0].as_i32() == args[1].as_i32()));
            }
            if kind == 2 {
                assert_eq!(
                    a.as_f64(),
                    f64::from(args[0].as_i32() == 0 || args[1].as_i32() != 0)
                );
            }
            checks += 1;
        }
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = chunks(fast, kind, 1);
            let body = &cs[1];
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            let validation =
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&bytes)
                    .map(|_| "valid".to_owned())
                    .unwrap_or_else(|e| e.message().to_owned());
            if kind < 3 {
                assert_eq!(validation, "valid");
                if std::env::var_os("VYBE_CONDITION_EXPORT").is_some() {
                    let mut export_cs = cs.clone();
                    for init in &cs[0].global_inits {
                        if let vybe_runtime::chunk::ConstExpr::RefFunc(index) = init.init {
                            let mut factory = Chunk::new(format!("binding-factory:{}", init.name));
                            factory.emit_op_u16(vybe_runtime::Op::REF_FUNC, index as u16, 0);
                            factory.emit(0, 0);
                            factory.emit_op(vybe_runtime::Op::RETURN, 0);
                            export_cs.push(factory);
                        }
                    }
                    native_export::export_chunks(&export_cs, &format!("condition-{kind}-{fast}"));
                }
            }
            sizes.push(serde_json::json!({"body":body.code.len(),"locals":body.local_count,"native_bytes":bytes.len(),"validation":validation}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        assert_eq!(sizes[0]["validation"], sizes[1]["validation"]);
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    let result = serde_json::json!({"cases":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../CONDITION_GATES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for kind in 0..4 {
        let mut ws = [worker(false, kind, 64), worker(true, kind, 64)];
        let args = if kind == 0 {
            vec![Value::I32(17)]
        } else if kind == 3 {
            vec![array((0..32).map(Value::I32).collect()), Value::I32(-1)]
        } else {
            vec![Value::I32(1), Value::I32(0)]
        };
        let run = |w: &mut (VM, Value)| {
            std::hint::black_box(w.0.invoke(&w.1, &args).unwrap());
        };
        for w in &mut ws {
            for _ in 0..8 {
                run(w);
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(200) || n < 8 {
                    run(&mut ws[side]);
                    n += 1;
                }
                samples[side].push(t.elapsed().as_secs_f64() / n as f64);
            }
        }
        let mut r: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        r.sort_by(f64::total_cmp);
        rows.push(
            serde_json::json!({"kind":kind,"median_speedup":r[3],"seconds_per_call":samples}),
        );
    }
    let result = serde_json::json!({"pairs":7,"minimum_sample_ms":200,"minimum_calls":8,"warmups":8,"operations_per_call":64,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../CONDITION_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
