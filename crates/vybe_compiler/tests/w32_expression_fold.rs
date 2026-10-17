#![cfg(feature = "primitive-benchmarks")]
use std::sync::{Arc, Mutex};
use vybe_ast::*;
use vybe_compiler::primitives::bits;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
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
fn output(n: &str, v: &serde_json::Value) {
    let p = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::write(
        p.join(format!("W32_EXPR_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn expression(op: IntOp, a: i64, b: i64) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Integer(op, BitLane::W32),
        left: Box::new(Expression::int(a)),
        right: Box::new(Expression::int(b)),
    })
}
fn compile(fast: bool, body: Vec<Statement>) -> Vec<Chunk> {
    bits::benchmark_with_word_expression_folding(fast, || {
        let mut profile =
            vybe_compiler::profile::parse_profile("[compiler]\ncase_sensitive=true\n").unwrap();
        profile.materialize_bool_results = true;
        profile.builtins.insert(
            "record".into(),
            vybe_runtime::profile::BuiltinDef {
                emit: vybe_runtime::profile::BuiltinEmit::HostCall(
                    "test:word-fold".into(),
                    "record".into(),
                ),
                discarded_emit: None,
                min_args: 1,
                max_args: 1,
                i32_condition: false,
                slot: None,
            },
        );
        let module = Module {
            name: "literal-word-fold".into(),
            language: Lang::Unknown,
            imports: vec![],
            canon: Default::default(),
            directives: Default::default(),
            body: vec![
                Statement::new(StmtKind::FunctionDecl {
                    name: "entry".into(),
                    params: vec![],
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
        let mut chunks = vybe_compiler::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        // Export genuine helper closures for the standard embedding aliases.
        let aliases: Vec<_> = chunks[0]
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
            vybe_compiler::primitives::functions::emit_ref_func(&mut f, i, 0, 0);
            f.emit_op(Op::RETURN, 0);
            chunks.push(f);
        }
        chunks
    })
}
fn worker(c: Vec<Chunk>) -> (VM, Value, Arc<Mutex<Vec<Value>>>) {
    let mut v = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut v);
    let log = Arc::new(Mutex::new(vec![]));
    let seen = log.clone();
    v.register_host_fn(
        "test:word-fold",
        "record",
        Box::new(move |_, a| {
            seen.lock().unwrap().push(a[0].clone());
            Value::Null
        }),
    );
    let f = v.run(c).unwrap();
    (v, f, log)
}
fn native(c: &[Chunk]) -> Vec<u8> {
    let b = vybe_platform_wasm::write_wasm_with_options(
        c,
        vybe_platform_wasm::WasmWriteOptions {
            include_name_section: false,
            include_vybe_metadata: false,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&b)
        .unwrap();
    b
}
fn expected(op: IntOp, a: i64, b: i64) -> Option<i32> {
    let (x, y) = (a as i32 as i128, b as i32 as i128);
    let (u, v) = (a as u32 as u128, b as u32 as u128);
    Some(match op {
        IntOp::Add => (x + y) as i32,
        IntOp::Sub => (x - y) as i32,
        IntOp::Mul => (x * y) as i32,
        IntOp::DivS => {
            if y == 0 || x == i32::MIN as i128 && y == -1 {
                return None;
            }
            (x / y) as i32
        }
        IntOp::DivU => {
            if v == 0 {
                return None;
            }
            (u / v) as i32
        }
        IntOp::RemS => {
            if y == 0 {
                return None;
            }
            (x % y) as i32
        }
        IntOp::RemU => {
            if v == 0 {
                return None;
            }
            (u % v) as i32
        }
        IntOp::Shl => (u << (v & 31)) as i32,
        IntOp::ShrS => (x >> (v & 31)) as i32,
        IntOp::ShrU => (u >> (v & 31)) as i32,
        IntOp::And => (x & y) as i32,
        IntOp::Or => (x | y) as i32,
        IntOp::Xor => (x ^ y) as i32,
        IntOp::Eq => i32::from(x == y),
        IntOp::Ne => i32::from(x != y),
        IntOp::LtS => i32::from(x < y),
        IntOp::LtU => i32::from(u < v),
        IntOp::LeS => i32::from(x <= y),
        IntOp::LeU => i32::from(u <= v),
        IntOp::GtS => i32::from(x > y),
        IntOp::GtU => i32::from(u > v),
        IntOp::GeS => i32::from(x >= y),
        IntOp::GeU => i32::from(u >= v),
    })
}
fn record(e: Expression) -> Statement {
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("record")),
        args: vec![Argument::positional(e)],
        optional: false,
    })))
}

fn tree(seed: usize, depth: usize) -> (Expression, i32, serde_json::Value) {
    let initial = [i32::MIN as i64, -1, 0, 1, i32::MAX as i64, u32::MAX as i64][seed % 6];
    let mut e = Expression::int(initial);
    let mut value = initial as i32;
    let mut steps = vec![];
    for k in 0..depth {
        let kind = (seed + k) % 28;
        if kind < 23 {
            let mut b = [-1, 0, 1, 31, 32, i32::MIN as i64, u32::MAX as i64][(seed * 3 + k) % 7];
            if expected(OPS[kind], value.into(), b).is_none() {
                b = 3;
            }
            value = expected(OPS[kind], value.into(), b).unwrap();
            e = Expression::new(ExprKind::Binary {
                op: BinOp::Integer(OPS[kind], BitLane::W32),
                left: Box::new(e),
                right: Box::new(Expression::int(b)),
            });
            steps.push(serde_json::json!([kind, b]));
        } else if kind < 25 {
            let b = [0, 1, 31, 32, -1, 65][(seed + k) % 6];
            let amount = (b as u32) & 31;
            // Independent 64-bit rotation formula, rather than rotate_*.
            let u = value as u32 as u64;
            value = if amount == 0 {
                value
            } else if kind == 23 {
                ((u << amount) | (u >> (32 - amount))) as i32
            } else {
                ((u >> amount) | (u << (32 - amount))) as i32
            };
            e = Expression::new(ExprKind::Binary {
                op: if kind == 23 {
                    BinOp::RotL(BitLane::W32)
                } else {
                    BinOp::RotR(BitLane::W32)
                },
                left: Box::new(e),
                right: Box::new(Expression::int(b)),
            });
            steps.push(serde_json::json!([kind, b]));
        } else {
            let bits: Vec<_> = (0..32).map(|i| ((value as u32 >> i) & 1) != 0).collect();
            value = match kind {
                25 => bits.iter().filter(|x| **x).count(),
                26 => bits.iter().rev().take_while(|x| !**x).count(),
                _ => bits.iter().take_while(|x| !**x).count(),
            } as i32;
            e = Expression::new(ExprKind::Unary {
                op: match kind {
                    25 => UnaryOp::PopCount(BitLane::W32),
                    26 => UnaryOp::LeadingZeros(BitLane::W32),
                    _ => UnaryOp::TrailingZeros(BitLane::W32),
                },
                expr: Box::new(e),
            });
            steps.push(serde_json::json!([kind, 0]));
        }
    }
    (
        e,
        value,
        serde_json::json!({"initial":initial,"steps":steps,"expected":value}),
    )
}
#[test]
fn mixed_constant_trees_match_independent_oracle() {
    let mut body = vec![];
    let mut want = vec![];
    let mut queries = vec![];
    for seed in 0..56 {
        for depth in [1, 4, 16, 40, 80] {
            let (e, v, q) = tree(seed, depth);
            body.push(record(e));
            want.push(Value::I32(v));
            queries.push(q);
        }
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::int(0)))));
    let old = compile(false, body.clone());
    let new = compile(true, body);
    let entry = |c: &[Chunk]| c.iter().find(|c| c.name == "entry").unwrap().code.len();
    let sizes = [entry(&old), entry(&new)];
    let native_sizes = [native(&old).len(), native(&new).len()];
    assert!(sizes[1] < sizes[0]);
    for (fast, c) in [(false, old), (true, new)] {
        native_export::export_chunks(
            &c,
            if fast {
                "word-expr-after"
            } else {
                "word-expr-before"
            },
        );
        let (mut v, f, log) = worker(c);
        v.invoke(&f, &[]).unwrap();
        assert_eq!(*log.lock().unwrap(), want);
    }
    output("CORPUS", &serde_json::json!({"queries":queries}));
    output(
        "GATES",
        &serde_json::json!({"checks":want.len()*2,"guest_bytes":sizes,"native_bytes":native_sizes,"baseline":"Previously verified literal-pair folding enabled on both sides"}),
    );
}
#[test]
fn nested_traps_and_effects_are_preserved() {
    let mut checks = 0;
    for op in [IntOp::DivS, IntOp::DivU, IntOp::RemS, IntOp::RemU] {
        let trap = Expression::new(ExprKind::Binary {
            op: BinOp::Integer(op, BitLane::W32),
            left: Box::new(tree(3, 8).0),
            right: Box::new(Expression::int(0)),
        });
        let outer = Expression::new(ExprKind::Binary {
            op: BinOp::Integer(IntOp::Mul, BitLane::W32),
            left: Box::new(trap),
            right: Box::new(Expression::int(0)),
        });
        let body = vec![
            record(Expression::int(77)),
            Statement::new(StmtKind::Return(Some(outer))),
        ];
        let mut results = vec![];
        for fast in [false, true] {
            let c = compile(fast, body.clone());
            native(&c);
            native_export::export_chunks(&c, &format!("word-expr-trap-{op:?}-{fast}"));
            let (mut v, f, log) = worker(c);
            let r = v.invoke(&f, &[]).map_err(|e| {
                (
                    e.message,
                    e.line,
                    e.call_stack
                        .into_iter()
                        .map(|frame| (frame.chunk_name, frame.line))
                        .collect::<Vec<_>>(),
                )
            });
            assert!(r.as_ref().unwrap_err().0.contains("integer divide by zero"));
            assert_eq!(*log.lock().unwrap(), vec![Value::F64(77.0)]);
            results.push(r);
        }
        assert_eq!(results[0], results[1]);
        checks += 2;
    }
    output("TRAPS", &serde_json::json!({"checks":checks}));
}
fn kernel(fast: bool, _op: IntOp, a: i64, _b: i64, repeats: usize) -> Vec<Chunk> {
    let depth = a as usize;
    let mut body: Vec<_> = (1..repeats)
        .map(|_| Statement::new(StmtKind::Expr(tree(7, depth).0)))
        .collect();
    body.push(Statement::new(StmtKind::Return(Some(tree(7, depth).0))));
    compile(fast, body)
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, op, a, b, control) in [
        ("tree4", IntOp::Add, 4, 0, false),
        ("tree16", IntOp::Add, 16, 0, false),
        ("tree40", IntOp::Add, 40, 0, false),
        ("control", IntOp::Add, 16, 0, true),
    ] {
        for repeats in [1, 16] {
            let (va, fa, _) = worker(kernel(false, op, a, b, repeats));
            let (vb, fb, _) = worker(kernel(!control, op, a, b, repeats));
            let mut vs = [va, vb];
            let fs = [fa, fb];
            let args: Vec<Value> = vec![];
            for side in 0..2 {
                assert_eq!(
                    vs[side].invoke(&fs[side], &args).unwrap(),
                    Value::I32(tree(7, a as usize).1)
                );
            }
            let mut wall = [vec![], vec![]];
            let mut cpu = [vec![], vec![]];
            let mut calls = vec![];
            for pair in 0..7 {
                let mut secs = [0., 0.];
                let mut cpus = [0., 0.];
                let mut counts = [0usize, 0];
                let mut turn = pair;
                while secs[0] < 0.1 || secs[1] < 0.1 || counts.iter().any(|n| *n < 8) {
                    for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                        if secs[side] >= 0.1 && counts[side] >= 8 {
                            continue;
                        }
                        let c = thread_cpu_seconds();
                        let t = Instant::now();
                        std::hint::black_box(vs[side].invoke(&fs[side], &args).unwrap());
                        secs[side] += t.elapsed().as_secs_f64();
                        cpus[side] += thread_cpu_seconds() - c;
                        counts[side] += 1;
                    }
                    turn += 1;
                }
                for side in 0..2 {
                    wall[side].push(secs[side] / counts[side] as f64);
                    cpu[side].push(cpus[side] / counts[side] as f64);
                }
                calls.push(counts);
            }
            let mut wr: Vec<_> = wall[0].iter().zip(&wall[1]).map(|(a, b)| a / b).collect();
            wr.sort_by(f64::total_cmp);
            let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(a, b)| a / b).collect();
            cr.sort_by(f64::total_cmp);
            rows.push(serde_json::json!({"label":label,"repeats":repeats,"median_speedup":wr[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
        }
    }
    output(
        "MEASUREMENTS",
        &serde_json::json!({"pairs":7,"minimum_ms_per_side":100,"separate_vms":true,"cases":rows}),
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

#[test]
fn side_effects_in_operand_sequences_are_preserved() {
    let effect = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("record")),
        args: vec![Argument::positional(Expression::int(77))],
        optional: false,
    });
    let e = Expression::new(ExprKind::Binary {
        op: BinOp::Integer(IntOp::Add, BitLane::W32),
        left: Box::new(Expression::new(ExprKind::Sequence(vec![
            effect,
            Expression::int(42),
        ]))),
        right: Box::new(tree(7, 16).0),
    });
    for fast in [false, true] {
        let c = compile(
            fast,
            vec![
                record(e.clone()),
                Statement::new(StmtKind::Return(Some(Expression::int(0)))),
            ],
        );
        native(&c);
        native_export::export_chunks(&c, &format!("word-expr-effect-{fast}"));
        let (mut v, f, log) = worker(c);
        v.invoke(&f, &[]).unwrap();
        assert_eq!(
            *log.lock().unwrap(),
            vec![
                Value::F64(77.0),
                Value::I32(42i32.wrapping_add(tree(7, 16).1))
            ]
        );
    }
    output(
        "EFFECT",
        &serde_json::json!({"expected":[77,42i32.wrapping_add(tree(7,16).1)],"checks":4}),
    );
}

#[test]
fn nested_word_condition_requests_preserve_truth() {
    let mut checks = 0;
    for seed in 0..28 {
        let (e, want, _) = tree(seed, 16);
        let body = vec![Statement::new(StmtKind::If {
            cond: e,
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::int(1))))],
            elifs: vec![],
            else_body: Some(vec![Statement::new(StmtKind::Return(Some(
                Expression::int(2),
            )))]),
        })];
        for fast in [false, true] {
            let c = compile(fast, body.clone());
            native(&c);
            let (mut v, f, _) = worker(c);
            assert_eq!(
                v.invoke(&f, &[]).unwrap().as_i32(),
                if want == 0 { 2 } else { 1 }
            );
            checks += 1;
        }
    }
    output("CONDITIONS", &serde_json::json!({"checks":checks}));
}

#[test]
fn compilation_cost_for_nonconstant_spines() {
    let mut rows = vec![];
    for (depth, control) in [(16, false), (64, false), (128, false), (64, true)] {
        let mut e = Expression::ident("input");
        for _ in 0..depth {
            e = Expression::new(ExprKind::Binary {
                op: BinOp::Integer(IntOp::Add, BitLane::W32),
                left: Box::new(e),
                right: Box::new(Expression::int(1)),
            });
        }
        let mut body: Vec<_> = (0..4)
            .map(|_| Statement::new(StmtKind::Expr(e.clone())))
            .collect();
        body.push(Statement::new(StmtKind::Return(Some(Expression::int(0)))));
        let before = compile(false, body.clone());
        let after = compile(!control, body.clone());
        for (a, b) in before.iter().zip(&after) {
            assert_eq!(a.code, b.code);
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let mut elapsed = 0.;
                let mut count = 0;
                while elapsed < 0.05 || count < 3 {
                    let input = body.clone();
                    let start = thread_cpu_seconds();
                    let result = compile(side == 1 && !control, input);
                    elapsed += thread_cpu_seconds() - start;
                    count += 1;
                    std::hint::black_box(&result);
                }
                samples[side].push(elapsed / count as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"depth":depth,"control":control,"median_thread_cpu_speedup":ratios[3],"seconds_per_compilation":samples}));
    }
    output(
        "COMPILE_COST",
        &serde_json::json!({"pairs":7,"minimum_thread_cpu_ms_per_side":50,"ast_cloning_outside_timing":true,"cases":rows}),
    );
}

#[test]
fn compact_and_dense_source_maps_match() {
    let body = vec![
        record(tree(11, 80).0),
        Statement::new(StmtKind::Return(Some(Expression::int(0)))),
    ];
    let dense = {
        let _scope = vybe_runtime::source_lines::LineEmissionScope::enter(false);
        compile(true, body.clone())
    };
    let compact = {
        let _scope = vybe_runtime::source_lines::LineEmissionScope::enter(true);
        compile(true, body)
    };
    for (a, b) in dense.iter().zip(&compact) {
        assert_eq!(a.code, b.code);
        assert_eq!(a.lines.len(), a.code.len());
        assert_eq!(b.lines.len(), b.code.len());
        assert!(a.lines.iter().eq(b.lines.iter()));
    }
    native(&dense);
    native(&compact);
    for c in [dense, compact] {
        let (mut v, f, log) = worker(c);
        v.invoke(&f, &[]).unwrap();
        assert_eq!(*log.lock().unwrap(), vec![Value::I32(tree(11, 80).1)]);
    }
    output(
        "SOURCE_MAPS",
        &serde_json::json!({"formats":2,"depth":80,"same_code_and_byte_indexed_lines":true}),
    );
}
