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
        p.join(format!("STATIC_WORD_{n}.json")),
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
fn compile_config(
    fast: bool,
    unsigned: bool,
    directives: Directives,
    body: Vec<Statement>,
    override_target: Option<(ProtocolSlot, &str)>,
) -> Vec<Chunk> {
    bits::benchmark_with_static_word_constants(fast, || {
        let mut profile = vybe_compiler::profile::parse_profile(
            "[compiler]\ncase_sensitive=true\n[builtin_types]\nint=[\"int\",\"integer\"]\n",
        )
        .unwrap();

        profile.materialize_bool_results = true;
        use vybe_ast::builtin_slots::BuiltinType;
        for (slot, target) in [
            (ProtocolSlot::Add, "opcode:i32.add"),
            (ProtocolSlot::Sub, "opcode:i32.sub"),
            (ProtocolSlot::Mul, "opcode:i32.mul"),
            (ProtocolSlot::Div, "opcode:i32.div_s"),
            (ProtocolSlot::IDiv, "opcode:i32.div_u"),
            (
                ProtocolSlot::Mod,
                if unsigned {
                    "opcode:i32.rem_u"
                } else {
                    "opcode:i32.rem_s"
                },
            ),
            (ProtocolSlot::And, "opcode:i32.and"),
            (ProtocolSlot::Or, "opcode:i32.or"),
            (ProtocolSlot::Xor, "opcode:i32.xor"),
            (ProtocolSlot::Eq, "opcode:i32.eq"),
            (ProtocolSlot::Ne, "opcode:i32.ne"),
            (
                ProtocolSlot::Lt,
                if unsigned {
                    "opcode:i32.lt_u"
                } else {
                    "opcode:i32.lt_s"
                },
            ),
            (
                ProtocolSlot::Le,
                if unsigned {
                    "opcode:i32.le_u"
                } else {
                    "opcode:i32.le_s"
                },
            ),
            (
                ProtocolSlot::Gt,
                if unsigned {
                    "opcode:i32.gt_u"
                } else {
                    "opcode:i32.gt_s"
                },
            ),
            (
                ProtocolSlot::Ge,
                if unsigned {
                    "opcode:i32.ge_u"
                } else {
                    "opcode:i32.ge_s"
                },
            ),
            (ProtocolSlot::LShift, "opcode:i32.shl"),
            (ProtocolSlot::RShift, "opcode:i32.shr_s"),
        ] {
            profile.builtin_slots.insert(BuiltinType::Int, slot, target);
        }

        if let Some((slot, target)) = override_target {
            profile.builtin_slots.insert(BuiltinType::Int, slot, target);
        }
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
            directives,
            body: vec![
                Statement::new(StmtKind::FunctionDecl {
                    name: "entry".into(),
                    params: vec![Param {
                        name: "a".into(),
                        type_hint: Some(TypeHint::descriptive("integer")),
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
fn compile_mode(
    fast: bool,
    unsigned: bool,
    directives: Directives,
    body: Vec<Statement>,
) -> Vec<Chunk> {
    compile_config(fast, unsigned, directives, body, None)
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

const KINDS: [usize; 15] = [0, 1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 15, 17, 19, 21];
fn generic(kind: usize) -> BinOp {
    match kind {
        0 => BinOp::Add,
        1 => BinOp::Sub,
        2 => BinOp::Mul,
        3 => BinOp::Div,
        4 => BinOp::IDiv,
        5 => BinOp::Mod,
        10 => BinOp::BitAnd,
        11 => BinOp::BitOr,
        12 => BinOp::BitXor,
        13 => BinOp::Eq,
        14 => BinOp::NotEq,
        15 => BinOp::Lt,
        17 => BinOp::LtEq,
        19 => BinOp::Gt,
        21 => BinOp::GtEq,
        _ => unreachable!(),
    }
}
fn live(kind: usize, b: i64, word: bool) -> Expression {
    let a = if word {
        Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::ident("a")),
            right: Box::new(Expression::int(0)),
        })
    } else {
        Expression::ident("a")
    };
    Expression::new(ExprKind::Binary {
        op: generic(kind),
        left: Box::new(a),
        right: Box::new(Expression::int(b)),
    })
}
fn directives() -> Directives {
    Directives {
        type_resolution: Some(TypeResolution::Static),
        operator_dispatch: Some(OperatorDispatch::StaticBuiltin),
        ..Default::default()
    }
}
fn compile(fast: bool, body: Vec<Statement>) -> Vec<Chunk> {
    compile_mode(fast, false, directives(), body)
}
fn effective(kind: usize, unsigned: bool) -> usize {
    if unsigned {
        match kind {
            5 => 6,
            15 => 16,
            17 => 18,
            19 => 20,
            21 => 22,
            _ => kind,
        }
    } else {
        kind
    }
}
#[test]
fn resolved_static_integer_operations_and_boolean_results() {
    let inputs = values();
    let mut checks = 0;
    let mut rows = vec![];
    for unsigned in [false, true] {
        let mut body = vec![];
        let mut queries = vec![];
        for word in [false, true] {
            for kind in KINDS {
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
                    if [3, 4, 5].contains(&kind) && b == 0 {
                        continue;
                    }
                    if kind == 3 && b == -1 {
                        continue;
                    }
                    body.push(record(live(kind, b, word)));
                    queries.push((kind, b, word));
                }
            }
        }
        body.push(Statement::new(StmtKind::Return(Some(Expression::int(0)))));
        let old = compile_mode(false, unsigned, directives(), body.clone());
        let new = compile_mode(true, unsigned, directives(), body);
        fn entry(c: &[Chunk]) -> &Chunk {
            c.iter().find(|c| c.name == "entry").unwrap()
        }
        let sizes = [entry(&old).code.len(), entry(&new).code.len()];
        let native_sizes = [native(&old).len(), native(&new).len()];
        assert!(sizes[1] < sizes[0]);
        for (fast, c) in [(false, old), (true, new)] {
            native_export::export_chunks(&c, &format!("static-word-{unsigned}-{fast}"));
            let (mut v, f, log) = worker(c);
            for a in &inputs {
                log.lock().unwrap().clear();
                v.invoke(&f, &[Value::F64(*a)]).unwrap();
                let want: Vec<_> = queries
                    .iter()
                    .map(|(k, b, _)| {
                        let n = oracle(effective(*k, unsigned), *a, *b);
                        if *k >= 13 {
                            Value::Bool(n != 0)
                        } else {
                            Value::I32(n)
                        }
                    })
                    .collect();
                assert_eq!(*log.lock().unwrap(), want, "unsigned={unsigned} a={a}");
                checks += want.len();
            }
        }
        rows.push(serde_json::json!({"unsigned":unsigned,"guest_bytes":sizes,"native_bytes":native_sizes,"queries":queries}));
    }
    output(
        "GATES",
        &serde_json::json!({"checks":checks,"rows":rows,"inputs":inputs}),
    );
}
#[test]
fn boolean_condition_requests_and_nested_word_consumers() {
    let mut checks = 0;
    for kind in [0, 2, 3, 4, 10, 11, 12, 13, 14, 15, 17, 19, 21] {
        let b = if [2, 3, 4].contains(&kind) { 1 } else { 0 };
        let condition = vec![Statement::new(StmtKind::If {
            cond: live(kind, b, true),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::int(1))))],
            elifs: vec![],
            else_body: Some(vec![Statement::new(StmtKind::Return(Some(
                Expression::int(2),
            )))]),
        })];
        for fast in [false, true] {
            let c = compile(fast, condition.clone());
            native(&c);
            let (mut v, f, _) = worker(c);
            for a in [-1., 0., 0.75, 1., u32::MAX as f64] {
                assert_eq!(
                    v.invoke(&f, &[Value::F64(a)]).unwrap().as_i32(),
                    if oracle(kind, a, b) == 0 { 2 } else { 1 }
                );
                checks += 1;
            }
        }
    }
    output("CONDITIONS", &serde_json::json!({"checks":checks}));
}
#[test]
fn static_shift_directives_keep_their_original_lowering() {
    let mut checks = 0;
    for policy in [ShiftOverflow::Mask, ShiftOverflow::Zero] {
        for op in [BinOp::Shl, BinOp::Shr] {
            for b in [-32, -1, 0, 31, 32, 64] {
                let mut d = directives();
                d.shift_overflow = Some(policy);
                let e = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(Expression::ident("a")),
                    right: Box::new(Expression::int(b)),
                });
                let body = vec![Statement::new(StmtKind::Return(Some(e)))];
                let old = compile_mode(false, false, d.clone(), body.clone());
                let new = compile_mode(true, false, d, body);
                for (a, b) in old.iter().zip(&new) {
                    assert_eq!(a.code, b.code);
                }
                let (mut a, fa, _) = worker(old);
                let (mut c, fc, _) = worker(new);
                for n in [-1., 0., 1., u32::MAX as f64] {
                    assert_eq!(
                        a.invoke(&fa, &[Value::F64(n)]).unwrap(),
                        c.invoke(&fc, &[Value::F64(n)]).unwrap()
                    );
                    checks += 1;
                }
            }
        }
    }
    output("SHIFT_GATES", &serde_json::json!({"checks":checks}));
}
fn kernel(fast: bool, _op: IntOp, kind: i64, b: i64, repeats: usize) -> Vec<Chunk> {
    let mut e = Expression::ident("a");
    for _ in 0..repeats {
        e = Expression::new(ExprKind::Binary {
            op: generic(kind as usize),
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
fn resolved_constant_pairs_and_trees() {
    let mut rows = vec![];
    let mut checks = 0;
    for unsigned in [false, true] {
        let mut body = vec![];
        let mut queries = vec![];
        for word in [false, true] {
            for kind in KINDS {
                for a in [i32::MIN as i64, -1, 0, 1, i32::MAX as i64, u32::MAX as i64] {
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
                        if expected(OPS[effective(kind, unsigned)], a, b).is_none() {
                            continue;
                        }
                        let left = if word {
                            Expression::new(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(Expression::int(a)),
                                right: Box::new(Expression::int(0)),
                            })
                        } else {
                            Expression::int(a)
                        };
                        let e = Expression::new(ExprKind::Binary {
                            op: generic(kind),
                            left: Box::new(left),
                            right: Box::new(Expression::int(b)),
                        });
                        body.push(record(e));
                        queries.push((kind, a, b, word));
                    }
                }
            }
        }
        body.push(Statement::new(StmtKind::Return(Some(Expression::int(0)))));
        let mut sizes = vec![];
        let mut native_sizes = vec![];
        for fast in [false, true] {
            let c = compile_mode(fast, unsigned, directives(), body.clone());
            sizes.push(c.iter().find(|c| c.name == "entry").unwrap().code.len());
            native_sizes.push(native(&c).len());
            native_export::export_chunks(&c, &format!("static-word-constant-{unsigned}-{fast}"));
            let (mut v, f, log) = worker(c);
            v.invoke(&f, &[Value::F64(0.)]).unwrap();
            let want: Vec<_> = queries
                .iter()
                .map(|(k, a, b, _)| {
                    let n = expected(OPS[effective(*k, unsigned)], *a, *b).unwrap();
                    if *k >= 13 {
                        Value::Bool(n != 0)
                    } else {
                        Value::I32(n)
                    }
                })
                .collect();
            assert_eq!(*log.lock().unwrap(), want);
            checks += want.len();
        }
        assert!(sizes[1] < sizes[0]);
        rows.push(serde_json::json!({"unsigned":unsigned,"guest_bytes":sizes,"native_bytes":native_sizes,"queries":queries}));
    }
    output(
        "CONSTANTS",
        &serde_json::json!({"checks":checks,"rows":rows}),
    );
}

#[test]
fn static_identity_conversion_and_arithmetic_traps() {
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
        (3, -1),
        (3, 0),
        (4, 0),
        (5, 0),
        (13, 0),
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
            native_export::export_chunks(&c, &format!("static-word-edge-{index}-{fast}"));
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

#[test]
fn dynamic_dispatch_and_nonword_target_keep_original_code() {
    let mut checks = 0;
    for runtime in [false, true] {
        let mut d = directives();
        if runtime {
            d.type_resolution = Some(TypeResolution::Dynamic);
            d.operator_dispatch = Some(OperatorDispatch::RuntimeProtocol);
        }
        let body = vec![Statement::new(StmtKind::Return(Some(live(0, 0, false))))];
        let old = compile_config(
            false,
            false,
            d.clone(),
            body.clone(),
            Some((ProtocolSlot::Add, "opcode:f64.add")),
        );
        let new = compile_config(
            true,
            false,
            d,
            body,
            Some((ProtocolSlot::Add, "opcode:f64.add")),
        );
        for (a, b) in old.iter().zip(&new) {
            assert_eq!(a.code, b.code);
        }
        for c in [old, new] {
            native(&c);
            let (mut v, f, _) = worker(c);
            for a in [-1., 0., 0.75, u32::MAX as f64] {
                assert_eq!(v.invoke(&f, &[Value::F64(a)]).unwrap(), Value::F64(a));
                checks += 1;
            }
        }
    }
    output("DISPATCH", &serde_json::json!({"checks":checks}));
}

fn infer_tree(seed: usize, depth: usize) -> Expression {
    let mut e = match seed % 7 {
        0 => Expression::int(1),
        1 => Expression::new(ExprKind::Lit(Literal::Float(1.25))),
        2 => Expression::new(ExprKind::Lit(Literal::BigInt(123))),
        3 => Expression::new(ExprKind::Lit(Literal::Bool(true))),
        4 => Expression::new(ExprKind::Lit(Literal::Str("x".into()))),
        5 => Expression::ident("unknown"),
        _ => Expression::new(ExprKind::Cast {
            type_name: "Alias".into(),
            expr: Box::new(Expression::int(1)),
        }),
    };
    for k in 0..depth {
        let right = if (seed + k) % 9 == 0 {
            Expression::new(ExprKind::Lit(Literal::BigInt(3)))
        } else {
            Expression::int(3)
        };
        e = Expression::new(ExprKind::Binary {
            op: [
                BinOp::Add,
                BinOp::Sub,
                BinOp::Mul,
                BinOp::Div,
                BinOp::IDiv,
                BinOp::FloorDiv,
                BinOp::Mod,
                BinOp::Pow,
                BinOp::BitAnd,
                BinOp::BitOr,
                BinOp::BitXor,
                BinOp::Shl,
                BinOp::Shr,
            ][(seed + k) % 13],
            left: Box::new(e),
            right: Box::new(right),
        });
    }
    e
}
#[test]
fn numeric_inference_matches_frozen_recursion() {
    let mut checks = 0;
    for config in [
        "[compiler]\n",
        "[compiler]\n[builtin_types]\nint=[\"Alias\"]\n",
        "[compiler]\n[builtin_types]\nbigint=[\"Alias\",\"int\"]\n",
        "[compiler]\n[builtin_types]\ndouble=[\"Alias\",\"int\"]\n",
    ] {
        let c = vybe_compiler::Compiler::with_profile(
            vybe_compiler::profile::parse_profile(config).unwrap(),
        );
        for seed in 0..91 {
            for depth in 0..9 {
                let e = infer_tree(seed, depth);
                let old = bits::benchmark_with_linear_numeric_inference(false, || {
                    c.benchmark_infer_expression_type(&e)
                });
                let new = bits::benchmark_with_linear_numeric_inference(true, || {
                    c.benchmark_infer_expression_type(&e)
                });
                assert_eq!(old, new, "seed={seed} depth={depth}");
                checks += 1;
            }
        }
    }
    output(
        "INFERENCE",
        &serde_json::json!({"checks":checks,"profiles":4,"frozen_recursion":true}),
    );
}

#[test]
fn numeric_inference_growth_measurements() {
    let c = vybe_compiler::Compiler::with_profile(
        vybe_compiler::profile::parse_profile("[compiler]\n").unwrap(),
    );
    let mut rows = vec![];
    for depth in [4, 8, 12, 16, 20] {
        let mut e = Expression::int(1);
        for _ in 0..depth {
            e = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(e),
                right: Box::new(Expression::int(1)),
            });
        }
        for fast in [false, true] {
            assert_eq!(
                bits::benchmark_with_linear_numeric_inference(fast, || c
                    .benchmark_infer_expression_type(&e)),
                Some("int".into())
            );
        }
        let mut samples = [vec![], vec![]];
        let mut counts = vec![];
        for pair in 0..7 {
            let mut pair_counts = [0, 0];
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                bits::benchmark_with_linear_numeric_inference(side == 1, || {
                    let mut elapsed = 0.;
                    let mut n = 0;
                    while elapsed < 0.025 || n < 1 {
                        let start = thread_cpu_seconds();
                        std::hint::black_box(
                            c.benchmark_infer_expression_type(std::hint::black_box(&e)),
                        );
                        elapsed += thread_cpu_seconds() - start;
                        n += 1;
                    }
                    samples[side].push(elapsed / n as f64);
                    pair_counts[side] = n;
                });
            }
            counts.push(pair_counts);
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"depth":depth,"median_thread_cpu_speedup":ratios[3],"seconds_per_query":samples,"counts":counts}));
    }
    output(
        "INFERENCE_MEASUREMENTS",
        &serde_json::json!({"pairs":7,"minimum_thread_cpu_ms_per_side":25,"scope":"Actual type inference, former recursion versus child-hint reuse; same classification rules","cases":rows}),
    );
}
#[test]
fn numeric_full_compilation_growth_measurements() {
    let mut rows = vec![];
    for depth in [4, 8, 12, 16] {
        let mut e = Expression::ident("a");
        for _ in 0..depth {
            e = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(e),
                right: Box::new(Expression::int(1)),
            });
        }
        let body = vec![Statement::new(StmtKind::Return(Some(e)))];
        let old =
            bits::benchmark_with_linear_numeric_inference(false, || compile(true, body.clone()));
        let new =
            bits::benchmark_with_linear_numeric_inference(true, || compile(true, body.clone()));
        for (a, b) in old.iter().zip(&new) {
            assert_eq!(a.code, b.code);
            assert_eq!(a.local_count, b.local_count);
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                bits::benchmark_with_linear_numeric_inference(side == 1, || {
                    let mut elapsed = 0.;
                    let mut n = 0;
                    while elapsed < 0.025 || n < 1 {
                        let input = body.clone();
                        let start = thread_cpu_seconds();
                        let result = compile(true, input);
                        elapsed += thread_cpu_seconds() - start;
                        n += 1;
                        std::hint::black_box(&result);
                    }
                    samples[side].push(elapsed / n as f64);
                });
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"depth":depth,"median_thread_cpu_speedup":ratios[3],"seconds_per_compile":samples,"emitted_code_and_local_count_identical":true}));
    }
    output(
        "COMPILE_GROWTH",
        &serde_json::json!({"pairs":7,"minimum_thread_cpu_ms_per_side":25,"ast_cloning_outside_timing":true,"scope":"Actual Compiler old recursive inference versus reused child hints; static optimizations enabled on both sides","cases":rows}),
    );
}

#[test]
fn arithmetic_opcode_in_comparison_slot_preserves_boolean_carrier() {
    let e = Expression::new(ExprKind::Binary {
        op: BinOp::Integer(IntOp::Add, BitLane::W32),
        left: Box::new(live(13, 0, false)),
        right: Box::new(Expression::int(1)),
    });
    let body = vec![Statement::new(StmtKind::Return(Some(e)))];
    let mut sides = vec![];
    for fast in [false, true] {
        let c = compile_config(
            fast,
            false,
            directives(),
            body.clone(),
            Some((ProtocolSlot::Eq, "opcode:i32.add")),
        );
        native(&c);
        let entry = c.iter().find(|c| c.name == "entry").unwrap();
        let dump = vybe_runtime::debug::disassemble(entry);
        assert_eq!(dump.matches("i64.trunc_f64_s").count(), 2, "{dump}");
        native_export::export_chunks(&c, &format!("static-word-carrier-{fast}"));
        let (mut v, f, _) = worker(c);
        let mut results = vec![];
        for a in [-1., 0., 0.75, 1., u32::MAX as f64] {
            results.push(v.invoke(&f, &[Value::F64(a)]).map_err(|e| e.message));
        }
        sides.push(results);
    }
    assert_eq!(sides[0], sides[1]);
    output(
        "CARRIER",
        &serde_json::json!({"checks":10,"mandatory_conversions":2}),
    );
}
