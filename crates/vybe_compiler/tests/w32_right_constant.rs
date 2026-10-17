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
        p.join(format!("W32_RIGHT_{n}.json")),
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
    bits::benchmark_with_word_right_constant(fast, || {
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
                    params: vec![Param {
                        name: "a".into(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    }],
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

fn live(kind: usize, b: i64, word: bool) -> Expression {
    let a = if word {
        Expression::new(ExprKind::Binary {
            op: BinOp::Integer(IntOp::Xor, BitLane::W32),
            left: Box::new(Expression::ident("a")),
            right: Box::new(Expression::int(0)),
        })
    } else {
        Expression::ident("a")
    };
    Expression::new(ExprKind::Binary {
        op: if kind < 23 {
            BinOp::Integer(OPS[kind], BitLane::W32)
        } else if kind == 23 {
            BinOp::RotL(BitLane::W32)
        } else {
            BinOp::RotR(BitLane::W32)
        },
        left: Box::new(a),
        right: Box::new(Expression::int(b)),
    })
}
fn oracle(kind: usize, a: f64, b: i64) -> i32 {
    let a = (a as i64) as i32;
    if kind < 23 {
        expected(OPS[kind], a.into(), b).unwrap()
    } else {
        let u = a as u32 as u64;
        let k = (b as u32) & 31;
        if k == 0 {
            a
        } else if kind == 23 {
            ((u << k) | (u >> (32 - k))) as i32
        } else {
            ((u >> k) | (u << (32 - k))) as i32
        }
    }
}
fn values() -> Vec<f64> {
    let mut v = vec![
        i32::MIN as f64 + 1.,
        -4294967295.,
        -65537.,
        -1.,
        -0.,
        0.,
        0.75,
        -0.75,
        1.,
        31.,
        i32::MAX as f64,
        u32::MAX as f64,
        4294967296.,
        9007199254740991.,
        -9007199254740991.,
    ];
    let mut n = 0x98765432u32;
    for _ in 0..64 {
        n = n.wrapping_mul(1664525).wrapping_add(1013904223);
        let x = n as i32;
        if x != i32::MIN {
            v.push(x as f64 + 0.25);
        }
    }
    v
}
#[test]
fn live_operands_all_word_operations_and_rotates() {
    let mut body = vec![];
    let mut queries = vec![];
    for word in [false, true] {
        for kind in 0..25 {
            for b in [
                i32::MIN as i64,
                -1,
                0,
                1,
                31,
                32,
                i32::MAX as i64,
                u32::MAX as i64,
            ] {
                if kind < 23
                    && matches!(
                        OPS[kind],
                        IntOp::DivS | IntOp::DivU | IntOp::RemS | IntOp::RemU
                    )
                    && b == 0
                {
                    continue;
                }
                // MIN / -1 is covered separately, without terminating this matrix.
                if kind == 3 && b == -1 {
                    continue;
                }
                body.push(record(live(kind, b, word)));
                queries.push((kind, b, word));
            }
        }
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::int(0)))));
    let old = compile(false, body.clone());
    let new = compile(true, body);
    fn entry(c: &[Chunk]) -> &Chunk {
        c.iter().find(|c| c.name == "entry").unwrap()
    }
    let sizes = [entry(&old).code.len(), entry(&new).code.len()];
    let locals = [entry(&old).local_count, entry(&new).local_count];
    let native_sizes = [native(&old).len(), native(&new).len()];
    assert!(sizes[1] < sizes[0]);
    let inputs = values();
    let mut checks = 0;
    for (fast, c) in [(false, old), (true, new)] {
        native_export::export_chunks(
            &c,
            if fast {
                "word-right-after"
            } else {
                "word-right-before"
            },
        );
        let (mut v, f, log) = worker(c);
        for a in &inputs {
            log.lock().unwrap().clear();
            v.invoke(&f, &[Value::F64(*a)]).unwrap();
            let want: Vec<_> = queries
                .iter()
                .map(|(k, b, _)| Value::I32(oracle(*k, *a, *b)))
                .collect();
            assert_eq!(*log.lock().unwrap(), want, "a={a}");
            checks += want.len();
        }
    }
    output(
        "CORPUS",
        &serde_json::json!({"queries":queries,"inputs":inputs}),
    );
    output(
        "GATES",
        &serde_json::json!({"checks":checks,"guest_bytes":sizes,"native_bytes":native_sizes,"locals":locals,"baseline":"All previous literal/tree folds enabled on both sides"}),
    );
}
#[test]
fn identity_conversion_traps_and_signed_overflow_stay_live() {
    let mut checks = 0;
    for (index, (kind, b)) in [
        (0, 0),
        (1, 0),
        (2, 1),
        (3, 1),
        (4, 1),
        (10, -1),
        (11, 0),
        (12, 0),
        (7, 32),
        (8, -32),
        (9, 0),
        (23, 64),
        (24, -32),
        (3, -1),
        (3, 0),
        (4, 0),
        (5, 0),
        (6, 0),
    ]
    .into_iter()
    .enumerate()
    {
        let body = vec![
            record(Expression::int(77)),
            Statement::new(StmtKind::Return(Some(live(kind, b, false)))),
        ];
        let mut sides = vec![];
        for fast in [false, true] {
            let c = compile(fast, body.clone());
            native(&c);
            native_export::export_chunks(&c, &format!("word-right-edge-{index}-{fast}"));
            let (mut v, f, log) = worker(c);
            let mut results = vec![];
            for a in [
                Value::F64(f64::NAN),
                Value::F64(f64::INFINITY),
                Value::F64(f64::NEG_INFINITY),
                Value::F64(9223372036854775808.),
                Value::F64(-9223372036854775808.),
                Value::F64(i32::MIN as f64),
                Value::F64(-0.),
                Value::Null,
                Value::Undefined,
                Value::Bool(true),
                Value::String("3".into()),
            ] {
                log.lock().unwrap().clear();
                let r = v.invoke(&f, &[a]).map_err(|e| {
                    (
                        e.message,
                        e.line,
                        e.call_stack
                            .into_iter()
                            .map(|f| (f.chunk_name, f.line))
                            .collect::<Vec<_>>(),
                    )
                });
                assert_eq!(*log.lock().unwrap(), vec![Value::F64(77.)]);
                results.push(r);
                checks += 1;
            }
            sides.push(results);
        }
        assert_eq!(sides[0], sides[1], "kind={kind} b={b}");
    }
    output("EDGE", &serde_json::json!({"checks":checks}));
}
fn kernel(fast: bool, _op: IntOp, kind: i64, b: i64, repeats: usize) -> Vec<Chunk> {
    let mut e = Expression::ident("a");
    for _ in 0..repeats {
        e = Expression::new(ExprKind::Binary {
            op: if kind < 23 {
                BinOp::Integer(OPS[kind as usize], BitLane::W32)
            } else {
                BinOp::RotL(BitLane::W32)
            },
            left: Box::new(e),
            right: Box::new(Expression::int(b)),
        });
    }
    compile(fast, vec![Statement::new(StmtKind::Return(Some(e)))])
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, op, a, b, control) in [
        ("live-add-zero", IntOp::Add, 0, 0, false),
        ("live-mul-one", IntOp::Add, 2, 1, false),
        ("live-div-one", IntOp::Add, 4, 1, false),
        ("live-rotate-zero", IntOp::Add, 23, 32, false),
        ("live-add-three", IntOp::Add, 0, 3, false),
        ("control", IntOp::Add, 0, 3, true),
    ] {
        for repeats in [1, 16, 128] {
            let (va, fa, _) = worker(kernel(false, op, a, b, repeats));
            let (vb, fb, _) = worker(kernel(!control, op, a, b, repeats));
            let mut vs = [va, vb];
            let fs = [fa, fb];
            let args = vec![Value::F64(1234567.0)];
            for side in 0..2 {
                assert_eq!(
                    vs[side].invoke(&fs[side], &args).unwrap(),
                    Value::I32((0..repeats).fold(1234567, |v, _| oracle(a as usize, v as f64, b)))
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
fn conditions_and_ordered_operand_effects() {
    let mut checks = 0;
    for (kind, b) in [
        (0, 0),
        (2, 1),
        (4, 1),
        (10, -1),
        (11, 0),
        (12, 0),
        (7, 32),
        (23, 64),
        (13, -1),
        (16, 0),
    ] {
        let body = vec![Statement::new(StmtKind::If {
            cond: live(kind, b, false),
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
            for a in [-1., -0., 1., i32::MAX as f64, u32::MAX as f64] {
                assert_eq!(
                    v.invoke(&f, &[Value::F64(a)]).unwrap().as_i32(),
                    if oracle(kind, a, b) == 0 { 2 } else { 1 }
                );
                checks += 1;
            }
        }
    }
    let effect = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("record")),
        args: vec![Argument::positional(Expression::int(77))],
        optional: false,
    });
    let left = Expression::new(ExprKind::Sequence(vec![effect, Expression::ident("a")]));
    let e = Expression::new(ExprKind::Binary {
        op: BinOp::Integer(IntOp::Add, BitLane::W32),
        left: Box::new(left),
        right: Box::new(Expression::int(0)),
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
        native_export::export_chunks(&c, &format!("word-right-effects-{fast}"));
        let (mut v, f, log) = worker(c);
        for a in [-1., 0., 0.75, u32::MAX as f64] {
            log.lock().unwrap().clear();
            v.invoke(&f, &[Value::F64(a)]).unwrap();
            assert_eq!(
                *log.lock().unwrap(),
                vec![Value::F64(77.), Value::I32(oracle(0, a, 0))]
            );
            checks += 2;
        }
    }
    output("REQUEST_EFFECTS", &serde_json::json!({"checks":checks}));
}

#[test]
fn compilation_cost_and_compact_source_maps() {
    let mut rows = vec![];
    for (label, depth, b, control) in [
        ("identity", 16, Some(0), false),
        ("identity", 128, Some(0), false),
        ("nonidentity", 128, Some(3), false),
        ("nonconstant", 128, None, false),
        ("control", 128, Some(0), true),
    ] {
        let mut e = Expression::ident("a");
        for _ in 0..depth {
            e = Expression::new(ExprKind::Binary {
                op: BinOp::Integer(IntOp::Add, BitLane::W32),
                left: Box::new(e),
                right: Box::new(
                    b.map(Expression::int)
                        .unwrap_or_else(|| Expression::ident("a")),
                ),
            });
        }
        let body = vec![Statement::new(StmtKind::Return(Some(e)))];
        let mut sizes = vec![];
        for fast in [false, true] {
            let dense = {
                let _s = vybe_runtime::source_lines::LineEmissionScope::enter(false);
                compile(fast && !control, body.clone())
            };
            let compact = {
                let _s = vybe_runtime::source_lines::LineEmissionScope::enter(true);
                compile(fast && !control, body.clone())
            };
            for (a, b) in dense.iter().zip(&compact) {
                assert_eq!(a.code, b.code);
                assert_eq!(a.lines.len(), a.code.len());
                assert_eq!(b.lines.len(), b.code.len());
                assert!(a.lines.iter().eq(b.lines.iter()));
            }
            sizes.push(dense.iter().find(|c| c.name == "entry").unwrap().code.len());
            for c in [dense, compact] {
                native(&c);
                let (mut v, f, _) = worker(c);
                assert_eq!(
                    v.invoke(&f, &[Value::F64(17.)]).unwrap(),
                    Value::I32(17 + depth * b.unwrap_or(17) as usize as i32)
                );
            }
        }
        let mut times = [vec![], vec![]];
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
                times[side].push(elapsed / count as f64);
            }
        }
        let mut ratios: Vec<_> = times[0].iter().zip(&times[1]).map(|(a, b)| a / b).collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"label":label,"depth":depth,"guest_bytes":sizes,"median_thread_cpu_speedup":ratios[3],"seconds_per_compilation":times}));
    }
    output(
        "COMPILE_COST",
        &serde_json::json!({"pairs":7,"minimum_thread_cpu_ms_per_side":50,"ast_cloning_outside_timing":true,"compact_dense_maps_equal":true,"cases":rows}),
    );
}
