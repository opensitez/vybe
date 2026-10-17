#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::*;
use vybe_runtime::{Chunk, Value, VM};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn compiler(fast: bool) -> vybe_compiler::Compiler {
    let profile = vybe_compiler::profile::parse_profile(
        "[compiler]\nnegative_index_wraps=true\ncase_sensitive=true\n",
    )
    .unwrap();
    let mut c = vybe_compiler::Compiler::with_profile(profile);
    c.benchmark_integer_truthiness_reuse(fast);
    c
}
fn chunks(fast: bool, kind: usize, repeats: usize) -> Vec<Chunk> {
    let e = Expression::new(ExprKind::Binary {
        op: [BinOp::Eqv, BinOp::Imp, BinOp::Xor][kind],
        left: Box::new(Expression::ident("a")),
        right: Box::new(Expression::ident("b")),
    });
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
    for kind in 0..3 {
        let mut ws = [worker(false, kind, 1), worker(true, kind, 1)];
        let values = [i32::MIN, -1024, -17, -2, 0, 1, 2, 17, i32::MAX];
        for a in values {
            for b in values {
                let args = [Value::I32(a), Value::I32(b)];
                let f0 = ws[0].1.clone();
                let f1 = ws[1].1.clone();
                let old = ws[0].0.invoke(&f0, &args).unwrap();
                let new = ws[1].0.invoke(&f1, &args).unwrap();
                let expected = match kind {
                    0 => a == b,
                    1 => a == 0 || b != 0,
                    _ => a != b,
                };
                assert_eq!(old, new, "kind={kind}, a={a}, b={b}");
                assert_eq!(new, Value::I32(i32::from(expected)));
                checks += 1;
            }
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
                if std::env::var_os("VYBE_INTEGER_EXPORT").is_some() {
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
                    native_export::export_chunks(&export_cs, &format!("integer-{kind}-{fast}"));
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
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../INTEGER_GATES.json"),
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
    for kind in 0..3 {
        let mut ws = [worker(false, kind, 64), worker(true, kind, 64)];
        let args = vec![Value::I32(17), Value::I32(4)];
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
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../INTEGER_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
