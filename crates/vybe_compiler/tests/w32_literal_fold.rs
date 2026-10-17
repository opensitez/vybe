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
        p.join(format!("W32_FOLD_{n}.json")),
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
    bits::benchmark_with_literal_word_folding(fast, || {
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
        let aliases:Vec<_>=chunks[0].global_inits.iter().filter_map(|g|
            if let vybe_runtime::chunk::ConstExpr::RefFunc(i)=g.init {Some((g.name.clone(),i))} else {None}).collect();
        for(name,i)in aliases {
            let mut f=Chunk::new(format!("binding-factory:{name}"));
            vybe_compiler::primitives::functions::emit_ref_func(&mut f,i,0,0);
            f.emit_op(Op::RETURN,0);chunks.push(f);
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
#[test]
fn all_operations_match_wide_integer_oracle_and_complete_modules() {
    let points = [
        i32::MIN as i64,
        -65537,
        -1,
        0,
        1,
        31,
        i32::MAX as i64,
        u32::MAX as i64,
    ];
    let mut body = vec![];
    let mut wants = vec![];
    let mut queries = vec![];
    for (oi, op) in OPS.into_iter().enumerate() {
        for a in points {
            for b in points {
                if let Some(w) = expected(op, a, b) {
                    body.push(record(expression(op, a, b)));
                    wants.push(Value::I32(w));
                    queries.push(serde_json::json!({"op":oi,"a":a,"b":b,"expected":w}));
                }
            }
        }
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::int(0)))));
    let old = compile(false, body.clone());
    let new = compile(true, body);
    let i = old.iter().position(|c| c.name == "entry").unwrap();
    let j = new.iter().position(|c| c.name == "entry").unwrap();
    let bytes = [native(&old), native(&new)];
    assert!(new[j].code.len() < old[i].code.len());
    for (fast, c) in [(false, old.clone()), (true, new.clone())] {
        native_export::export_chunks(
            &c,
            if fast {
                "word-fold-after"
            } else {
                "word-fold-before"
            },
        );
        let (mut v, f, log) = worker(c);
        v.invoke(&f, &[]).unwrap();
        assert_eq!(*log.lock().unwrap(), wants);
    }
    output(
        "GATES",
        &serde_json::json!({"checks":wants.len()*2,"old_body":old[i].code.len(),"new_body":new[j].code.len(),"old_native":bytes[0].len(),"new_native":bytes[1].len()}),
    );
    output("CORPUS", &serde_json::json!({"queries":queries}));
}
#[test]
fn traps_and_wider_literals_keep_original_emission() {
    let mut checks = 0;
    for op in OPS {
        for (a, b) in [
            (i64::MAX, 1),
            (i64::MIN, 1),
            (4294967296, 3),
            (-2147483649, 3),
            (9007199254740993, 1),
        ] {
            let body = vec![Statement::new(StmtKind::Return(Some(expression(op, a, b))))];
            let old = compile(false, body.clone());
            let new = compile(true, body);
            for (x, y) in old.iter().zip(&new) {
                assert_eq!(x.code, y.code);
                assert_eq!(x.local_count, y.local_count);
            }
            checks += 1;
        }
    }
    for (op, a, b) in [
        (IntOp::DivS, 1, 0),
        (IntOp::DivU, 1, 0),
        (IntOp::RemS, 1, 0),
        (IntOp::RemU, 1, 0),
        (IntOp::DivS, i32::MIN as i64, -1),
    ] {
        let body = vec![Statement::new(StmtKind::Return(Some(expression(op, a, b))))];
        let old = compile(false, body.clone());
        let new = compile(true, body);
        for (x, y) in old.iter().zip(&new) {
            assert_eq!(x.code, y.code);
        }
        let (mut a, fa, _) = worker(old);
        let (mut b, fb, _) = worker(new);
        let error = |e: vybe_runtime::VMError| e.to_string();
        assert_eq!(
            a.invoke(&fa, &[]).map_err(error),
            b.invoke(&fb, &[]).map_err(error)
        );
        checks += 1;
    }
    output("FALLBACK_GATES", &serde_json::json!({"checks":checks}));
}

fn kernel(fast: bool, op: IntOp, a: i64, b: i64, repeats: usize) -> Vec<Chunk> {
    let mut body: Vec<_> = (1..repeats)
        .map(|_| Statement::new(StmtKind::Expr(expression(op, a, b))))
        .collect();
    body.push(Statement::new(StmtKind::Return(Some(expression(op, a, b)))));
    compile(fast, body)
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, op, a, b, control) in [
        ("add", IntOp::Add, 2147483647, 1, false),
        ("unsigned-div", IntOp::DivU, 4294967295, 3, false),
        ("shift", IntOp::ShrU, -1, 17, false),
        ("comparison", IntOp::LtU, -1, 2147483647, false),
        ("control", IntOp::Add, 2147483647, 1, true),
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
                    Value::I32(expected(op, a, b).unwrap())
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
fn conditions_and_nested_word_consumers_preserve_representation() {
    let cases = [
        (IntOp::Eq, -1, 4294967295),
        (IntOp::LtS, -1, 1),
        (IntOp::Add, -2147483648, -2147483648),
        (IntOp::Or, 1, -2147483648),
        (IntOp::RemS, -2147483648, -1),
    ];
    let mut checks = 0;
    for (index, (op, a, b)) in cases.into_iter().enumerate() {
        let want = expected(op, a, b).unwrap();
        for fast in [false, true] {
            let conditional = vec![Statement::new(StmtKind::If {
                cond: expression(op, a, b),
                then_body: vec![Statement::new(StmtKind::Return(Some(Expression::int(1))))],
                elifs: vec![],
                else_body: Some(vec![Statement::new(StmtKind::Return(Some(
                    Expression::int(2),
                )))]),
            })];
            let c = compile(fast, conditional);
            native(&c);
            native_export::export_chunks(&c, &format!("word-fold-condition-{index}-{fast}"));
            let (mut v, f, _) = worker(c);
            assert_eq!(
                v.invoke(&f, &[]).unwrap().as_i32(),
                if want == 0 { 2 } else { 1 }
            );
            checks += 1;
            let nested = Expression::new(ExprKind::Binary {
                op: BinOp::Integer(IntOp::Xor, BitLane::W32),
                left: Box::new(expression(op, a, b)),
                right: Box::new(Expression::int(17)),
            });
            let c = compile(fast, vec![Statement::new(StmtKind::Return(Some(nested)))]);
            native(&c);
            let (mut v, f, _) = worker(c);
            assert_eq!(v.invoke(&f, &[]).unwrap(), Value::I32(want ^ 17));
            checks += 1;
        }
    }
    output("REQUEST_GATES", &serde_json::json!({"checks":checks}));
}
