#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::*;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn output(name: &str, v: &serde_json::Value) {
    let p = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&p).unwrap();
    std::fs::write(
        p.join(format!("WORD_REF_{name}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
const OPS: [IntOp; 23] = [
    IntOp::Add,
    IntOp::Sub,
    IntOp::Mul,
    IntOp::DivS,
    IntOp::DivU,
    IntOp::RemS,
    IntOp::RemU,
    IntOp::Shl,
    IntOp::ShrS,
    IntOp::ShrU,
    IntOp::And,
    IntOp::Or,
    IntOp::Xor,
    IntOp::Eq,
    IntOp::Ne,
    IntOp::LtS,
    IntOp::LtU,
    IntOp::LeS,
    IntOp::LeU,
    IntOp::GtS,
    IntOp::GtU,
    IntOp::GeS,
    IntOp::GeU,
];
fn operation(k: usize, lane: BitLane) -> Expression {
    if k >= 100 {
        let mut expression = operation(k % 100, lane);
        match &mut expression.kind {
            ExprKind::Binary { op, left, right } => {
                *right = Box::new(Expression::int(
                    if matches!(
                        op,
                        BinOp::Integer(IntOp::Or | IntOp::Shl | IntOp::ShrS | IntOp::ShrU, _)
                            | BinOp::RotL(_)
                            | BinOp::RotR(_)
                    ) {
                        0
                    } else {
                        1
                    },
                ));
                if k >= 200 {
                    *left = Box::new(Expression::new(ExprKind::Binary {
                        op: BinOp::Integer(IntOp::Add, BitLane::W32),
                        left: Box::new(Expression::ident("a")),
                        right: Box::new(Expression::int(0)),
                    }));
                }
            }
            ExprKind::Unary { expr, .. } if k >= 200 => {
                *expr = Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Integer(IntOp::Add, BitLane::W32),
                    left: Box::new(Expression::ident("a")),
                    right: Box::new(Expression::int(0)),
                }));
            }
            _ => {}
        }
        return expression;
    }
    let a = Box::new(Expression::ident("a"));
    let b = Box::new(Expression::ident("b"));
    Expression::new(if k < 23 {
        ExprKind::Binary {
            op: BinOp::Integer(OPS[k], lane),
            left: a,
            right: b,
        }
    } else if k < 25 {
        ExprKind::Binary {
            op: if k == 23 {
                BinOp::RotL(lane)
            } else {
                BinOp::RotR(lane)
            },
            left: a,
            right: b,
        }
    } else {
        ExprKind::Unary {
            op: match k {
                25 => UnaryOp::PopCount(lane),
                26 => UnaryOp::LeadingZeros(lane),
                _ => UnaryOp::TrailingZeros(lane),
            },
            expr: a,
        }
    })
}
fn ret(n: i64) -> Statement {
    Statement::new(StmtKind::Return(Some(Expression::int(n))))
}
fn condition(e: Expression) -> Statement {
    Statement::new(StmtKind::If {
        cond: e,
        then_body: vec![ret(1)],
        elifs: vec![],
        else_body: Some(vec![ret(0)]),
    })
}
fn compile(fast: bool, lane: BitLane, only: Option<usize>, repeats: usize) -> Vec<Chunk> {
    compile_mode(fast, lane, only, repeats, false)
}
fn compile_mode(
    fast: bool,
    lane: BitLane,
    only: Option<usize>,
    repeats: usize,
    protocol: bool,
) -> Vec<Chunk> {
    vybe_compiler::primitives::bits::benchmark_with_word_scalar_reference_proof(fast, || {
        let profile =
            vybe_compiler::profile::parse_profile("[compiler]\ncase_sensitive=true\n").unwrap();
        let mut body = vec![];
        if only == Some(99) {
            body.push(condition(Expression::ident("a")));
        } else if let Some(k) = only {
            for _ in 0..repeats.saturating_sub(1) {
                body.push(Statement::new(StmtKind::If {
                    cond: operation(k, lane),
                    then_body: vec![],
                    elifs: vec![],
                    else_body: None,
                }));
            }
            body.push(condition(operation(k, lane)));
        } else {
            for k in 0..28 {
                body.push(Statement::new(StmtKind::If {
                    cond: Expression::new(ExprKind::Binary {
                        op: BinOp::Integer(IntOp::Eq, BitLane::W32),
                        left: Box::new(Expression::ident("kind")),
                        right: Box::new(Expression::int(k as i64)),
                    }),
                    then_body: vec![condition(operation(k, lane))],
                    elifs: vec![],
                    else_body: None,
                }));
            }
            body.push(ret(-1));
        }
        let params = ["a", "b", "kind"]
            .into_iter()
            .map(|name| Param {
                name: name.into(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect();
        let m = Module {
            name: "word-ref".into(),
            language: Lang::Unknown,
            imports: vec![],
            canon: Default::default(),
            directives: Directives {
                truthiness: protocol.then_some(Truthiness::Protocol),
                ..Default::default()
            },
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
        };
        vybe_compiler::Compiler::with_profile(profile)
            .compile(&m)
            .unwrap()
    })
}
fn worker(c: Vec<Chunk>) -> (VM, Value) {
    let mut v = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut v);
    let f = v.run(c).unwrap();
    (v, f)
}
fn native(c: &[Chunk]) -> Vec<u8> {
    let b = vybe_platform_wasm::write_wasm_with_options(
        c,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: false,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&b)
        .unwrap();
    b
}
#[test]
fn results_traps_and_native_modules() {
    let vals = [
        i32::MIN as f64,
        -65537.,
        -1.,
        -0.,
        0.,
        0.75,
        1.,
        32.,
        65537.,
        i32::MAX as f64,
        u32::MAX as f64,
        f64::NAN,
        f64::INFINITY,
    ];
    let mut checks = 0;
    let mut sizes = vec![];
    let mut queries = vec![];
    for (lane_id, lane) in [BitLane::W32, BitLane::W64].into_iter().enumerate() {
        let old = compile(false, lane, None, 1);
        let new = compile(true, lane, None, 1);
        for (fast, cs) in [(false, &old), (true, &new)] {
            let bytes = native(cs);
            let entry = cs.iter().find(|x| x.name == "entry").unwrap();
            sizes.push(serde_json::json!({"lane":lane_id,"fast":fast,"guest":entry.code.len(),"native":bytes.len(),"locals":entry.local_count}));
            native_export::export_chunks(cs, &format!("word-ref-{lane_id}-{fast}"));
        }
        let (mut a, fa) = worker(old);
        let (mut b, fb) = worker(new);
        for kind in 0..28 {
            for &x in &vals {
                for &y in &vals {
                    let args = [Value::F64(x), Value::F64(y), Value::F64(kind as f64)];
                    let old = a
                        .invoke(&fa, &args)
                        .map(|v| format!("{v:?}"))
                        .map_err(|e| e.message);
                    let new = b
                        .invoke(&fb, &args)
                        .map(|v| format!("{v:?}"))
                        .map_err(|e| e.message);
                    assert_eq!(old, new, "lane={lane:?} kind={kind} x={x:?} y={y:?}");
                    checks += 1;
                    if lane_id == 0 {
                        queries.push((
                            format!("{:016x}", x.to_bits()),
                            format!("{:016x}", y.to_bits()),
                            kind,
                        ));
                    }
                }
            }
        }
        for kind in [0, 10, 23, 25, 26, 27] {
            for bad in [
                Value::Null,
                Value::Undefined,
                Value::Bool(true),
                Value::String("1".into()),
                Value::I64(i64::MAX),
            ] {
                let args = [bad, Value::F64(1.), Value::F64(kind as f64)];
                assert_eq!(
                    a.invoke(&fa, &args)
                        .map(|v| format!("{v:?}"))
                        .map_err(|e| e.message),
                    b.invoke(&fb, &args)
                        .map(|v| format!("{v:?}"))
                        .map_err(|e| e.message)
                );
                checks += 1;
            }
        }
    }
    output("GATES", &serde_json::json!({"checks":checks,"sizes":sizes}));
    output("QUERIES", &serde_json::json!({"queries":queries}));
}
#[test]
fn protocol_numeric_conditions_preserve_results_and_traps() {
    let mut checks = 0;
    for (lane_id, lane) in [BitLane::W32, BitLane::W64].into_iter().enumerate() {
        let old = compile_mode(false, lane, None, 1, true);
        let new = compile_mode(true, lane, None, 1, true);
        for (fast, cs) in [(false, &old), (true, &new)] {
            native(cs);
            native_export::export_chunks(cs, &format!("word-ref-protocol-{lane_id}-{fast}"));
        }
        let (mut a, fa) = worker(old);
        let (mut b, fb) = worker(new);
        for kind in 0..28 {
            for x in [
                -1.,
                -0.,
                0.,
                1.,
                31.,
                i32::MAX as f64,
                f64::NAN,
                f64::INFINITY,
            ] {
                for y in [0., 1., -1., 32.] {
                    let args = [Value::F64(x), Value::F64(y), Value::F64(kind as f64)];
                    assert_eq!(
                        a.invoke(&fa, &args)
                            .map(|v| format!("{v:?}"))
                            .map_err(|e| e.message),
                        b.invoke(&fb, &args)
                            .map(|v| format!("{v:?}"))
                            .map_err(|e| e.message),
                        "kind={kind} lane={lane:?} x={x:?} y={y:?}"
                    );
                    checks += 1;
                }
            }
        }
    }
    output("PROTOCOL", &serde_json::json!({"checks":checks}));
}
#[test]
fn unknown_value_and_existing_comparisons_keep_exact_code() {
    for k in [13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 99] {
        for lane in [BitLane::W32, BitLane::W64] {
            let a = compile(false, lane, Some(k), 1);
            let b = compile(true, lane, Some(k), 1);
            assert_eq!(a.len(), b.len());
            for (x, y) in a.iter().zip(&b) {
                assert_eq!(x.code, y.code, "kind={k} lane={lane:?} chunk={}", x.name);
            }
        }
    }
}
#[test]
fn literal_operands_and_nested_physical_words() {
    let mut checks = 0;
    for lane in [BitLane::W32, BitLane::W64] {
        for base in [100, 200] {
            for kind in 0..28 {
                let (mut a, fa) = worker(compile(false, lane, Some(base + kind), 1));
                let (mut b, fb) = worker(compile(true, lane, Some(base + kind), 1));
                for x in [
                    -2147483648.,
                    -1.,
                    -0.,
                    0.,
                    1.,
                    2147483647.,
                    4294967295.,
                    f64::NAN,
                    f64::INFINITY,
                ] {
                    let args = [Value::F64(x), Value::F64(1.), Value::F64(0.)];
                    assert_eq!(
                        a.invoke(&fa, &args)
                            .map(|v| format!("{v:?}"))
                            .map_err(|e| e.message),
                        b.invoke(&fb, &args)
                            .map(|v| format!("{v:?}"))
                            .map_err(|e| e.message),
                        "lane={lane:?} base={base} kind={kind} x={x:?}"
                    );
                    checks += 1;
                }
            }
        }
    }
    output("LITERALS", &serde_json::json!({"checks":checks}));
}
#[test]
#[ignore]
fn debug_pairs() {
    let mut rows = vec![];
    for lane in [BitLane::W32, BitLane::W64] {
        for kind in [0, 10, 13, 23, 25, 26, 27] {
            for repeats in [1, 16] {
                let (va, fa) = worker(compile(false, lane, Some(kind), repeats));
                let (vb, fb) = worker(compile(true, lane, Some(kind), repeats));
                let mut vs = [va, vb];
                let fs = [fa, fb];
                let args = [Value::F64(7.), Value::F64(1.), Value::F64(kind as f64)];
                assert_eq!(
                    vs[0].invoke(&fs[0], &args).unwrap(),
                    vs[1].invoke(&fs[1], &args).unwrap()
                );
                let mut samples = [vec![], vec![]];
                let mut cpu = [vec![], vec![]];
                for pair in 0..7 {
                    let mut seconds = [0., 0.];
                    let mut cpus = [0., 0.];
                    let mut calls = [0usize, 0];
                    let mut turn = pair;
                    while seconds.iter().any(|s| *s < 0.1) {
                        for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                            if seconds[side] >= 0.1 {
                                continue;
                            }
                            let c = thread_cpu_seconds();
                            let t = std::time::Instant::now();
                            std::hint::black_box(vs[side].invoke(&fs[side], &args).unwrap());
                            seconds[side] += t.elapsed().as_secs_f64();
                            cpus[side] += thread_cpu_seconds() - c;
                            calls[side] += 1;
                        }
                        turn += 1;
                    }
                    for side in 0..2 {
                        samples[side].push(seconds[side] / calls[side] as f64);
                        cpu[side].push(cpus[side] / calls[side] as f64);
                    }
                }
                let mut ratios: Vec<_> = samples[0]
                    .iter()
                    .zip(&samples[1])
                    .map(|(x, y)| x / y)
                    .collect();
                ratios.sort_by(f64::total_cmp);
                let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(x, y)| x / y).collect();
                cr.sort_by(f64::total_cmp);
                rows.push(serde_json::json!({"lane":format!("{lane:?}"),"kind":kind,"repeats":repeats,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"samples":samples,"cpu":cpu}));
            }
        }
    }
    output(
        "MEASUREMENTS",
        &serde_json::json!({"pairs":7,"minimum_ms_per_side":100,"cases":rows}),
    );
}
#[cfg(target_os = "macos")]
fn thread_cpu_seconds() -> f64 {
    #[repr(C)]
    struct Timespec {
        sec: i64,
        nsec: i64,
    }
    unsafe extern "C" {
        fn clock_gettime(clock: i32, result: *mut Timespec) -> i32;
    }
    let mut t = Timespec { sec: 0, nsec: 0 };
    assert_eq!(unsafe { clock_gettime(16, &mut t) }, 0);
    t.sec as f64 + t.nsec as f64 * 1e-9
}
#[cfg(not(target_os = "macos"))]
fn thread_cpu_seconds() -> f64 {
    0.0
}
