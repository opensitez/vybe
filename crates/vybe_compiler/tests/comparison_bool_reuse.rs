#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::*;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn operator(i: usize) -> BinOp {
    [
        BinOp::Eq,
        BinOp::NotEq,
        BinOp::StrictEq,
        BinOp::StrictNotEq,
        BinOp::Lt,
        BinOp::LtEq,
        BinOp::Gt,
        BinOp::GtEq,
        BinOp::In,
        BinOp::NotIn,
    ][i]
}
fn compare(i: usize) -> Expression {
    Expression::new(ExprKind::Binary {
        op: operator(i),
        left: Box::new(Expression::ident("a")),
        right: Box::new(Expression::ident("b")),
    })
}
fn program(fast: bool, op: usize, mode: u8, repeats: usize) -> Vec<Chunk> {
    let expr = if mode == 4 {
        Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(compare(op)),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        })
    } else {
        compare(op)
    };
    let params = ["a", "b"]
        .into_iter()
        .map(|name| Param {
            name: name.into(),
            type_hint: matches!(mode, 0 | 5 | 6).then(|| TypeHint::checked("double")),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        })
        .collect();
    let body = if mode == 5 {
        vec![Statement::new(StmtKind::If {
            cond: expr,
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::int(1))))],
            elifs: vec![],
            else_body: Some(vec![Statement::new(StmtKind::Return(Some(Expression::int(0))))]),
        })]
    } else {
        let mut body: Vec<_> = (1..repeats)
            .map(|_| Statement::new(StmtKind::Expr(expr.clone())))
            .collect();
        body.push(Statement::new(StmtKind::Return(Some(expr))));
        body
    };
    let module = Module {
        name: "boolean-reuse".into(),
        language: Lang::CSharp,
        imports: vec![],
        canon: Default::default(),
        directives: Directives {
            type_resolution: Some(TypeResolution::Static),
            operator_dispatch: Some(OperatorDispatch::Ecma),
            ecma_operator_coercion: Some(true),
            comparison_results_are_boolean: (mode == 3).then_some(false),
            ..Default::default()
        },
        body: vec![
            Statement::new(StmtKind::FunctionDecl {
                name: "compare".into(),
                params,
                body,
                return_type: None,
                modifiers: Modifiers::default(),
                handles: vec![],
                is_async: false,
                is_generator: false,
                is_sub: false,
            }),
            Statement::new(StmtKind::Return(Some(Expression::ident("compare")))),
        ],
    };
    let mut profile = vybe_compiler::profile::parse_profile(&format!(
        "[compiler]\nmaterialize_bool_results = {}\n",
        mode != 2
    ))
    .unwrap();
    if mode == 6 {
        // The existing constructor for a custom numeric recipe must survive;
        // only the redundant conversion after that Boolean may disappear.
        profile.builtin_slots.insert(
            vybe_ast::builtin_slots::BuiltinType::Double,
            ProtocolSlot::Lt,
            "opcode:f64.add",
        );
    }
    let mut compiler = vybe_compiler::Compiler::with_profile(profile);
    compiler.benchmark_comparison_bool_reuse(fast);
    let mut cs = compiler.compile(&module).unwrap();
    let aliases: Vec<_> = cs[0]
        .global_inits
        .iter()
        .filter_map(|g| {
            if let vybe_runtime::chunk::ConstExpr::RefFunc(i) = g.init {
                Some((g.name.clone(), i))
            } else {
                None
            }
        })
        .collect();
    for (name, i) in aliases {
        let mut f = Chunk::new(format!("binding-factory:{name}"));
        f.emit_op_u16(Op::REF_FUNC, i as u16, 0);
        f.emit(0, 0);
        f.emit_op(Op::RETURN, 0);
        cs.push(f);
    }
    cs
}
fn worker(fast: bool, op: usize, mode: u8, repeats: usize) -> (VM, Value) {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, op, mode, repeats)).unwrap();
    (vm, f)
}
fn object(array: bool) -> Value {
    let o = if array {
        vybe_runtime::value::Object::new_array(vec![Value::I32(1), Value::Null])
    } else {
        let mut o = vybe_runtime::value::Object::new();
        o.properties.insert("key".into(), Value::I32(7));
        o
    };
    Value::Object(vybe_runtime::heap::alloc(o))
}
#[test]
fn boolean_reuse_preserves_tags_values_and_unproven_paths() {
    let mut checks = 0;
    let values = [
        Value::F64(-0.0),
        Value::I32(0),
        Value::I32(1),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::String("key".into()),
        Value::String("".into()),
        Value::String("é😀".into()),
        Value::Bool(true),
        Value::Null,
        Value::Undefined,
        Value::bigint_i64(1),
        object(false),
        object(true),
    ];
    for op in 0..10 {
        for mode in [1, 2, 3, 4] {
            let mut workers = [worker(false, op, mode, 1), worker(true, op, mode, 1)];
            for a in &values {
                for b in &values {
                    let args = [a.clone(), b.clone()];
                    let out = workers
                        .each_mut()
                        .map(|(vm, f)| vm.invoke(f, &args).map_err(|e| e.message));
                    assert_eq!(out[0], out[1], "op={op} mode={mode} a={a:?} b={b:?}");
                    if let Ok(v) = &out[1] {
                        if mode != 2 {
                            assert!(matches!(v, Value::Bool(_)), "{v:?}");
                        }
                    }
                    checks += 1;
                }
            }
        }
    }
    for op in 0..8 {
        let mut workers = [worker(false, op, 0, 1), worker(true, op, 0, 1)];
        for a in [-0.0, 0.25, 17.0, -17.0, f64::NAN, f64::INFINITY] {
            for b in [0.0, 4.0, -3.0, f64::NAN, f64::NEG_INFINITY] {
                let expected = [a == b, a != b, a == b, a != b, a < b, a <= b, a > b, a >= b][op];
                for (vm, f) in &mut workers {
                    assert_eq!(
                        vm.invoke(f, &[Value::F64(a), Value::F64(b)]).unwrap(),
                        Value::Bool(expected)
                    );
                }
                checks += 1;
            }
        }
    }
    for mode in [2, 3, 5] {
        let old = program(false, 4, mode, 1);
        let new = program(true, 4, mode, 1);
        assert_eq!(
            old.iter().map(|c| &c.code).collect::<Vec<_>>(),
            new.iter().map(|c| &c.code).collect::<Vec<_>>(),
            "mode={mode}"
        );
    }
    for fast in [false, true] {
        let (mut vm, f) = worker(fast, 4, 6, 1);
        for (a, b) in [(3.0, -3.0), (3.0, 7.0), (f64::NAN, 1.0)] {
            let sum = a + b;
            assert_eq!(vm.invoke(&f, &[Value::F64(a), Value::F64(b)]).unwrap(),
                // Preserve the existing fromI32 producer's VM contract,
                // including NaN supplied by this deliberately custom recipe.
                Value::Bool(sum != 0.0));
            checks += 1;
        }
    }
    for fast in [false, true] {
        for op in [8, 9] {
            let (mut vm, f) = worker(fast, op, 1, 1);
            let o = object(false);
            let key = format!("{o}");
            if let Value::Object(handle) = &o {
                handle.lock().unwrap().properties.insert(key, Value::I32(7));
            }
            assert_eq!(vm.invoke(&f, &[o.clone(), o]).unwrap(), Value::Bool(op == 8));
            checks += 1;
        }
    }
    eprintln!(
        "Verified {checks} comparison Boolean type/value cases; numeric/raw-condition policy controls unchanged"
    );
}
#[test]
fn size_and_native_exports() {
    let mut rows = vec![];
    for mode in [0u8, 1, 2, 3, 4, 5, 6] {
        for op in 0..if mode < 2 {
            if mode == 0 { 8 } else { 10 }
        } else {
            1
        } {
            let op = if mode >= 2 { 4 } else { op };
            let mut sizes = vec![];
            for fast in [false, true] {
                let cs = program(fast, op, mode, 1);
                let c = cs.iter().find(|c| c.name == "compare").unwrap();
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
                if std::env::var_os("VYBE_PRIMITIVE_BOOL_REUSE_EXPORT").is_some() {
                    native_export::export_chunks(
                        &cs,
                        &format!("bool-reuse-{mode}-{op}-{}", u8::from(fast)),
                    );
                }
                sizes.push(serde_json::json!({"body":c.code.len(),"native":bytes.len(),"locals":c.local_count}));
            }
            if mode == 0 || mode == 4 {
                assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
            }
            rows.push(serde_json::json!({"mode":mode,"op":op,"sizes":sizes}));
        }
    }
    if std::env::var_os("VYBE_PRIMITIVE_BOOL_REUSE_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../BOOL_REUSE_GATES.json"),
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
    for (label, mode, a, b) in [
        ("typed-numbers", 0, Value::F64(3.0), Value::F64(7.0)),
        ("dynamic-numbers", 1, Value::F64(3.0), Value::F64(7.0)),
        (
            "dynamic-strings",
            1,
            Value::String("abc".into()),
            Value::String("def".into()),
        ),
        ("nested", 4, Value::F64(3.0), Value::F64(7.0)),
        (
            "numeric-policy-control",
            3,
            Value::F64(3.0),
            Value::F64(7.0),
        ),
        ("raw-condition-control", 5, Value::F64(3.0), Value::F64(7.0)),
    ] {
        let mut workers = [worker(false, 4, mode, 64), worker(true, 4, mode, 64)];
        let args = [a, b];
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, f) = &mut workers[side];
                let start = Instant::now();
                let mut n = 0;
                while start.elapsed() < Duration::from_millis(100) || n < 8 {
                    std::hint::black_box(vm.invoke(f, &args).unwrap());
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
        rows.push(
            serde_json::json!({"case":label,"median_speedup":ratios[3],"seconds_per_call":samples}),
        );
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"minimum_calls_per_sample":8,"comparisons_per_call":64,"condition_control_comparisons_per_call":1,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../BOOL_REUSE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
