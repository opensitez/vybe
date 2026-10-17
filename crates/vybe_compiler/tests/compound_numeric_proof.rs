#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::*;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn binary(op: BinOp, a: Expression, b: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(a),
        right: Box::new(b),
    })
}
fn expression(depth: usize, sequence: bool) -> Expression {
    // A lowering sequence keeps a checked numeric tail, but intentionally
    // has no inferred nominal type. Assignment makes evaluation observable.
    let input = if sequence {
        Expression::new(ExprKind::Sequence(vec![
            Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("a")),
                value: Box::new(Expression::new(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(Expression::ident("a")),
                })),
            }),
            Expression::ident("a"),
        ]))
    } else {
        Expression::ident("a")
    };
    let mut v = binary(BinOp::Add, input, Expression::ident("b"));
    for _ in 0..depth {
        v = binary(
            BinOp::Mul,
            v,
            binary(BinOp::Sub, Expression::ident("b"), Expression::ident("a")),
        );
    }
    binary(BinOp::Lt, v, Expression::ident("c"))
}
fn program(fast: bool, checked: bool, depth: usize, repeats: usize) -> Vec<Chunk> {
    program_variant(fast, checked, depth, repeats, 0)
}
fn program_variant(
    fast: bool,
    checked: bool,
    depth: usize,
    repeats: usize,
    custom_binding: u8,
) -> Vec<Chunk> {
    let hint = if custom_binding == 3 {
        TypeHint::converting("double")
    } else if checked {
        TypeHint::checked("double")
    } else {
        TypeHint::descriptive("double")
    };
    let params = ["a", "b", "c"]
        .into_iter()
        .map(|name| Param {
            name: name.into(),
            type_hint: Some(hint.clone()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        })
        .collect();
    let mut body: Vec<_> = (1..repeats)
        .map(|_| Statement::new(StmtKind::Expr(expression(depth, custom_binding != 4))))
        .collect();
    body.push(Statement::new(StmtKind::Return(Some(expression(
        depth,
        custom_binding != 4,
    )))));
    let module = Module {
        name: "compound-numeric".into(),
        language: Lang::CSharp,
        imports: vec![],
        canon: Default::default(),
        directives: Directives {
            type_resolution: Some(TypeResolution::Static),
            operator_dispatch: Some(OperatorDispatch::Ecma),
            ecma_operator_coercion: Some(true),
            add_concatenates_strings: Some(true),
            ..Default::default()
        },
        body: vec![
            Statement::new(StmtKind::FunctionDecl {
                name: "calculate".into(),
                params,
                body,
                return_type: None,
                modifiers: Modifiers::default(),
                handles: vec![],
                is_async: false,
                is_generator: false,
                is_sub: false,
            }),
            Statement::new(StmtKind::Return(Some(Expression::ident("calculate")))),
        ],
    };
    let mut profile =
        vybe_compiler::profile::parse_profile("[compiler]\nmaterialize_bool_results = true\n")
            .unwrap();
    if custom_binding == 1 {
        profile.builtin_slots.insert(
            vybe_ast::builtin_slots::BuiltinType::Double,
            ProtocolSlot::Sub,
            "opcode:f64.eq",
        );
    }
    if custom_binding == 2 {
        profile.builtin_slots.insert(
            vybe_ast::builtin_slots::BuiltinType::Double,
            ProtocolSlot::Sub,
            "opcode:i32.sub",
        );
    }
    let mut compiler = vybe_compiler::Compiler::with_profile(profile);
    compiler.benchmark_compound_numeric_proof(fast);
    let mut cs = compiler.compile(&module).unwrap();
    let bindings: Vec<_> = cs[0]
        .global_inits
        .iter()
        .filter_map(|init| {
            if let vybe_runtime::chunk::ConstExpr::RefFunc(index) = init.init {
                Some((init.name.clone(), index))
            } else {
                None
            }
        })
        .collect();
    for (name, index) in bindings {
        let mut f = Chunk::new(format!("binding-factory:{name}"));
        f.emit_op_u16(Op::REF_FUNC, index as u16, 0);
        f.emit(0, 0);
        f.emit_op(Op::RETURN, 0);
        cs.push(f);
    }
    cs
}
fn worker(fast: bool, checked: bool, depth: usize, repeats: usize) -> (VM, Value) {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, checked, depth, repeats)).unwrap();
    (vm, f)
}
#[test]
fn compound_numeric_proof_keeps_values_and_advisory_hints_dynamic() {
    let mut checks = 0;
    for depth in [0, 1, 4] {
        let mut workers = [worker(false, true, depth, 1), worker(true, true, depth, 1)];
        for a in [
            -0.0,
            0.25,
            3.0,
            -17.0,
            f64::NAN,
            f64::INFINITY,
            f64::NEG_INFINITY,
        ] {
            for b in [0.0, 4.0, -3.0, f64::NAN, f64::INFINITY] {
                for c in [-1.0, 0.0, 7.0, f64::INFINITY] {
                    let args = [Value::F64(a), Value::F64(b), Value::F64(c)];
                    let out = workers
                        .each_mut()
                        .map(|(vm, f)| vm.invoke(f, &args).unwrap());
                    assert_eq!(out[0], out[1]);
                    let a = -a;
                    let mut v = a + b;
                    for _ in 0..depth {
                        v *= b - a;
                    }
                    assert_eq!(out[1], Value::Bool(v < c));
                    checks += 1;
                }
            }
        }
    }
    for checked in [false] {
        let old = program(false, checked, 2, 1);
        let new = program(true, checked, 2, 1);
        assert_eq!(
            old.iter().map(|c| &c.code).collect::<Vec<_>>(),
            new.iter().map(|c| &c.code).collect::<Vec<_>>()
        );
        let mut workers = [worker(false, checked, 2, 1), worker(true, checked, 2, 1)];
        for a in [
            Value::String("3".into()),
            Value::Null,
            Value::Bool(true),
            Value::bigint_i64(2),
        ] {
            let args = [a, Value::F64(4.0), Value::F64(7.0)];
            let out = workers
                .each_mut()
                .map(|(vm, f)| vm.invoke(f, &args).map_err(|e| e.message));
            assert_eq!(out[0], out[1]);
            checks += 1;
        }
    }
    eprintln!("Verified {checks} compound arithmetic proof cases");
}
#[test]
fn size_and_native_exports() {
    let mut rows = vec![];
    for depth in [0, 1, 4] {
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, true, depth, 1);
            let function = cs.iter().find(|c| c.name == "calculate").unwrap();
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
            if std::env::var_os("VYBE_PRIMITIVE_COMPOUND_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("compound-{depth}-{}", u8::from(fast)));
            }
            sizes.push(serde_json::json!({"body":function.code.len(),"locals":function.local_count,"native":bytes.len(),"all_bodies":cs.iter().map(|c|c.code.len()).sum::<usize>()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        rows.push(serde_json::json!({"depth":depth,"sizes":sizes}));
    }
    if std::env::var_os("VYBE_PRIMITIVE_COMPOUND_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../COMPOUND_GATES.json"),
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
    for depth in [0, 1, 4] {
        let mut workers = [
            worker(false, true, depth, 32),
            worker(true, true, depth, 32),
        ];
        let args = [Value::F64(3.0), Value::F64(4.0), Value::F64(8.0)];
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
        rows.push(serde_json::json!({"depth":depth,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"minimum_calls_per_sample":8,"expressions_per_call":32,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../COMPOUND_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}

#[test]
fn arbitrary_builtin_slot_result_does_not_inherit_argument_number_proof() {
    let old = program_variant(false, true, 1, 1, 1);
    let new = program_variant(true, true, 1, 1, 1);
    assert_eq!(
        old.iter().map(|c| &c.code).collect::<Vec<_>>(),
        new.iter().map(|c| &c.code).collect::<Vec<_>>()
    );
}

#[test]
fn checked_word_arithmetic_result_keeps_wrapping_and_numeric_domain() {
    let programs = [
        program_variant(false, true, 1, 1, 2),
        program_variant(true, true, 1, 1, 2),
    ];
    let mut checks = 0;
    for cs in &programs {
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        let f = vm.run(cs.clone()).unwrap();
        for a in [i32::MIN, -17, 0, 3, i32::MAX] {
            for b in [i32::MIN, -3, 0, 4, i32::MAX] {
                for c in [-1.0, 0.0, 7.0, f64::INFINITY] {
                    let expected = (-(a as f64) + b as f64) * (b.wrapping_add(a) as f64) < c;
                    let out = vm
                        .invoke(
                            &f,
                            &[Value::F64(a as f64), Value::F64(b as f64), Value::F64(c)],
                        )
                        .unwrap();
                    assert_eq!(out, Value::Bool(expected));
                    checks += 1;
                }
            }
        }
    }
    for (side, cs) in programs.iter().enumerate() {
        let bytes = vybe_platform_wasm::write_wasm(cs);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
        if std::env::var_os("VYBE_PRIMITIVE_COMPOUND_EXPORT").is_some() {
            native_export::export_chunks(cs, &format!("compound-word-{side}"));
        }
    }
    eprintln!("Verified {checks} checked word arithmetic/overflow controls");
}

#[test]
fn straight_line_typed_control_is_already_direct() {
    let old = program_variant(false, true, 4, 1, 4);
    let new = program_variant(true, true, 4, 1, 4);
    assert_eq!(
        old.iter().map(|c| &c.code).collect::<Vec<_>>(),
        new.iter().map(|c| &c.code).collect::<Vec<_>>()
    );
}
#[test]
fn converting_declarations_preserve_numeric_entry_coercion() {
    let programs = [
        program_variant(false, true, 1, 1, 3),
        program_variant(true, true, 1, 1, 3),
    ];
    for (side, cs) in programs.iter().enumerate() {
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        let f = vm.run(cs.clone()).unwrap();
        for args in [
            [
                Value::String("3".into()),
                Value::String("4".into()),
                Value::String("8".into()),
            ],
            [Value::I32(3), Value::I32(4), Value::I32(8)],
            [Value::F64(3.0), Value::F64(4.0), Value::F64(8.0)],
        ] {
            assert_eq!(vm.invoke(&f, &args).unwrap(), Value::Bool(true));
        }
        let bytes = vybe_platform_wasm::write_wasm(cs);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
        if std::env::var_os("VYBE_PRIMITIVE_COMPOUND_EXPORT").is_some() {
            native_export::export_chunks(cs, &format!("compound-converting-{side}"));
        }
    }
    eprintln!("Verified 6 converting declaration controls");
}
