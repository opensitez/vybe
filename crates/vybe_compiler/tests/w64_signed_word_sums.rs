#![cfg(feature = "primitive-benchmarks")]
use vybe_ast::*;
use vybe_compiler::primitives::bits;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/w64_signed_word_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
const OPS: [IntOp; 2] = [IntOp::Add, IntOp::Sub];
fn output(n: &str, v: &serde_json::Value) {
    let p = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::write(
        p.join(format!("W64_SUMS_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn native(cs: &[Chunk]) -> Vec<u8> {
    let b = vybe_platform_wasm::write_wasm_with_options(
        cs,
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
fn worker(cs: Vec<Chunk>) -> (VM, Value) {
    let mut v = VM::new();
    v.register_host_fn("test:wide-words", "read", Box::new(|_, a| a[0].clone()));
    let f = v.run(cs).unwrap();
    (v, f)
}
fn kernel(fast: bool, op: IntOp, words: (bool, bool), repeats: usize, prefix: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("signed-wide-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    if prefix {
        c.emit_i32_const(23, 0);
    }
    for i in 0..repeats {
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        if fast {
            bits::emit_wide_sum_with_signed_word_operands(&mut c, op, words, 0);
        } else {
            before::emit_integer_binary(&mut c, BitLane::W64, op, 0);
        }
        if i + 1 < repeats || prefix {
            c.emit_op(Op::DROP, 0);
        }
    }
    c.emit_op(Op::RETURN, 0);
    vec![root, c]
}
fn bin(op: BinOp, a: Expression, b: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(a),
        right: Box::new(b),
    })
}
fn read(name: &str) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("observe")),
        args: vec![
            Argument::positional(Expression::ident(name)),
            Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(name.into())))),
        ],
        optional: false,
    })
}
fn operand(kind: usize, name: &str) -> Expression {
    let id = Expression::ident(name);
    match kind {
        0 => id,
        1 => bin(
            BinOp::Integer(IntOp::Or, BitLane::W32),
            id,
            Expression::int(0),
        ),
        2 => bin(BinOp::RotL(BitLane::W32), id, Expression::int(3)),
        3 => Expression::new(ExprKind::Unary {
            op: UnaryOp::PopCount(BitLane::W32),
            expr: Box::new(id),
        }),
        4 => Expression::int(if name == "a" { -7 } else { 3 }),
        5 => bin(
            BinOp::Integer(IntOp::Or, BitLane::W32),
            read(name),
            Expression::int(0),
        ),
        6 => read(name),
        _ => unreachable!(),
    }
}
fn compiled_expr(
    fast: bool,
    op: IntOp,
    a: Expression,
    b: Expression,
    repeats: usize,
    exact: bool,
) -> Vec<Chunk> {
    bits::benchmark_with_wide_signed_words(fast, || {
        let mut profile =
            vybe_compiler::profile::parse_profile("[compiler]\ncase_sensitive=true\n").unwrap();
        profile
            .builtin_return_types
            .insert("observe".into(), "integer".into());
        profile.builtins.insert(
            "observe".into(),
            vybe_runtime::profile::BuiltinDef {
                emit: vybe_runtime::profile::BuiltinEmit::HostCall(
                    "test:wide-words".into(),
                    "read".into(),
                ),
                discarded_emit: None,
                min_args: 2,
                max_args: 2,
                i32_condition: false,
                slot: None,
            },
        );
        let e = bin(BinOp::Integer(op, BitLane::W64), a, b);
        let mut body: Vec<_> = (1..repeats)
            .map(|_| Statement::new(StmtKind::Expr(e.clone())))
            .collect();
        body.push(Statement::new(StmtKind::Return(Some(e))));
        let module = Module {
            name: "signed-wide-compiler".into(),
            language: Lang::Unknown,
            imports: vec![],
            canon: Default::default(),
            directives: Directives {
                wide_integer_transport: if exact {
                    Some(WideIntegerTransport::Exact)
                } else {
                    None
                },
                ..Default::default()
            },
            body: vec![
                Statement::new(StmtKind::FunctionDecl {
                    name: "entry".into(),
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
                Statement::new(StmtKind::Return(Some(Expression::ident("entry")))),
            ],
        };
        vybe_compiler::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap()
    })
}
fn compiled(fast: bool, op: IntOp, kinds: (usize, usize), repeats: usize) -> Vec<Chunk> {
    compiled_expr(
        fast,
        op,
        operand(kinds.0, "a"),
        operand(kinds.1, "b"),
        repeats,
        false,
    )
}
fn recipe(kind: usize, name: &str, x: f64) -> f64 {
    match kind {
        0 | 6 => x,
        1 | 5 => ((x as i64) as i32) as f64,
        2 => ((x as i64) as i32).rotate_left(3) as f64,
        3 => ((x as i64) as i32).count_ones() as f64,
        4 => {
            if name == "a" {
                -7.
            } else {
                3.
            }
        }
        _ => unreachable!(),
    }
}
fn expected(op: IntOp, a: f64, b: f64) -> f64 {
    let (a, b) = (a as i64, b as i64);
    (match op {
        IntOp::Add => a + b,
        IntOp::Sub => a - b,

        _ => unreachable!(),
    }) as f64
}
fn check_result(value: Value, expected: f64) {
    assert!(
        matches!(value, Value::F64(_)),
        "W64 arithmetic must retain F64: {value:?}"
    );
    assert_eq!(value.as_f64().to_bits(), expected.to_bits());
}
#[test]
fn kernels_match_frozen_arithmetic_and_independent_oracle() {
    let points = [i32::MIN, -65537, -1, 0, 1, 17, 65536, i32::MAX];
    let mut pairs: Vec<_> = points
        .into_iter()
        .flat_map(|a| points.into_iter().map(move |b| (a, b)))
        .collect();
    let mut seed = 0x73628194u32;
    let mut next = || {
        seed ^= seed << 13;
        seed ^= seed >> 17;
        seed ^= seed << 5;
        seed as i32
    };
    for _ in 0..4096 {
        pairs.push((next(), next()));
    }
    let mut checks = 0;
    let mut sizes = vec![];
    for (oi, op) in OPS.into_iter().enumerate() {
        for (mask, words) in [(false, false), (true, false), (false, true), (true, true)]
            .into_iter()
            .enumerate()
        {
            let old = kernel(false, op, words, 1, false);
            let new = kernel(true, op, words, 1, false);
            let bytes = [native(&old), native(&new)];
            if mask != 3 {
                assert_eq!(old[1].code, new[1].code);
                assert_eq!(old[1].local_count, new[1].local_count);
            } else {
                assert!(new[1].code.len() < old[1].code.len());
            }
            sizes.push(serde_json::json!({"op":oi,"mask":mask,"old_body":old[1].code.len(),"new_body":new[1].code.len(),"old_locals":old[1].local_count,"new_locals":new[1].local_count,"old_native":bytes[0].len(),"new_native":bytes[1].len()}));
            native_export::export_chunks(&old, &format!("sums-kernel-{oi}-{mask}-before"));
            native_export::export_chunks(&new, &format!("sums-kernel-{oi}-{mask}-after"));
            let (mut a, fa) = worker(old);
            let (mut b, fb) = worker(new);
            for &(x, y) in &pairs {
                let args = [
                    if words.0 {
                        Value::I32(x)
                    } else {
                        Value::F64(x as f64)
                    },
                    if words.1 {
                        Value::I32(y)
                    } else {
                        Value::F64(y as f64)
                    },
                ];
                let want = expected(op, x as f64, y as f64);
                check_result(a.invoke(&fa, &args).unwrap(), want);
                check_result(b.invoke(&fb, &args).unwrap(), want);
                checks += 2;
            }
            for fast in [false, true] {
                let (mut v, f) = worker(kernel(fast, op, words, 1, true));
                assert_eq!(
                    v.invoke(&f, &[Value::I32(-7), Value::I32(3)]).unwrap(),
                    Value::I32(23)
                );
                checks += 1;
            }
        }
    }
    output(
        "KERNEL_GATES",
        &serde_json::json!({"checks":checks,"sizes":sizes}),
    );
}
#[test]
fn compiler_recipes_literals_and_effects() {
    let kinds = [
        (1, 1),
        (1, 0),
        (0, 1),
        (2, 3),
        (4, 1),
        (5, 5),
        (6, 0),
        (0, 0),
        (4, 4),
    ];
    let points = [
        -2147483648.,
        -65537.,
        -1.,
        0.,
        1.,
        17.,
        65536.,
        2147483647.,
        2147483648.,
        4294967295.,
    ];
    let mut checks = 0;
    let mut sizes = vec![];
    let mut queries = vec![];
    for (oi, op) in OPS.into_iter().enumerate() {
        for (case, &kind) in kinds.iter().enumerate() {
            let old = compiled(false, op, kind, 1);
            let new = compiled(true, op, kind, 1);
            let idx = old.iter().position(|c| c.name == "entry").unwrap();
            let ni = new.iter().position(|c| c.name == "entry").unwrap();
            let nb = [native(&old), native(&new)];
            if [1, 2, 6, 7].contains(&case) {
                assert_eq!(old[idx].code, new[ni].code);
                assert_eq!(old[idx].local_count, new[ni].local_count);
            } else {
                assert!(new[ni].code.len() < old[idx].code.len(), "case {case}");
            }
            sizes.push(serde_json::json!({"op":oi,"case":case,"kind":kind,"old_body":old[idx].code.len(),"new_body":new[ni].code.len(),"old_native":nb[0].len(),"new_native":nb[1].len()}));
            native_export::export_chunks(&old, &format!("sums-compiled-{oi}-{case}-before"));
            native_export::export_chunks(&new, &format!("sums-compiled-{oi}-{case}-after"));
            let (mut a, fa) = worker(old);
            let (mut b, fb) = worker(new);
            for x in points {
                for y in points {
                    let want = expected(op, recipe(kind.0, "a", x), recipe(kind.1, "b", y));
                    let args = [Value::F64(x), Value::F64(y)];
                    check_result(a.invoke(&fa, &args).unwrap(), want);
                    check_result(b.invoke(&fb, &args).unwrap(), want);
                    checks += 2;
                    queries
                        .push(serde_json::json!({"op":oi,"case":case,"a":x,"b":y,"expected":want}));
                }
            }
            for args in [
                [Value::F64(f64::NAN), Value::F64(1.)],
                [Value::F64(1.), Value::F64(f64::INFINITY)],
                [Value::F64(9223372036854775808.), Value::F64(-1.)],
                [Value::Null, Value::I32(1)],
            ] {
                // Smaller bytecode changes stack offsets, not the trap kind.
                let trap =
                    |e: vybe_runtime::VMError| e.to_string().lines().next().unwrap().to_owned();
                let one = a.invoke(&fa, &args).map_err(trap);
                let two = b.invoke(&fb, &args).map_err(trap);
                assert_eq!(one, two);
                checks += 1;
            }
        }
    }
    let literal_pairs = [
        (-7, 0),
        (i32::MIN as i64, -1),
        (i32::MAX as i64, -1),
        (-1, 4294967295),
        (4294967295, -1),
        (0, 4294967295),
    ];
    for (oi, op) in OPS.into_iter().enumerate() {
        for (case, &(a, b)) in literal_pairs.iter().enumerate() {
            let mut sides = vec![];
            for fast in [false, true] {
                let cs = compiled_expr(fast, op, Expression::int(a), Expression::int(b), 1, false);
                native(&cs);
                native_export::export_chunks(&cs, &format!("sums-literal-{oi}-{case}-{fast}"));
                let (mut v, f) = worker(cs.clone());
                check_result(
                    v.invoke(&f, &[Value::Null, Value::Null]).unwrap(),
                    expected(op, a as f64, b as f64),
                );
                checks += 1;
                sides.push(cs);
            }
            if a > i32::MAX as i64 || b > i32::MAX as i64 {
                for (x, y) in sides[0].iter().zip(&sides[1]) {
                    assert_eq!(x.code, y.code);
                }
            }
        }
    }
    for op in OPS {
        let a = compiled_expr(
            false,
            op,
            Expression::int(i64::MIN),
            Expression::int(i64::MAX),
            1,
            true,
        );
        let b = compiled_expr(
            true,
            op,
            Expression::int(i64::MIN),
            Expression::int(i64::MAX),
            1,
            true,
        );
        native(&a);
        native(&b);
        for (x, y) in a.iter().zip(&b) {
            assert_eq!(x.code, y.code);
            assert_eq!(x.local_count, y.local_count);
        }
    }
    output(
        "COMPILER_GATES",
        &serde_json::json!({"checks":checks,"sizes":sizes}),
    );
    output("CORPUS", &serde_json::json!({"queries":queries}));
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, shape, control) in [
        ("both-words", 0, false),
        ("compiler-word-pair", 3, false),
        ("compiler-literals", 4, false),
        ("compiler-mixed", 5, false),
        ("control", 0, true),
    ] {
        for repeats in [1, 16] {
            let make = |fast| match shape {
                0 => kernel(fast, IntOp::Add, (true, true), repeats, false),
                1 => kernel(fast, IntOp::Add, (true, false), repeats, false),
                2 => kernel(fast, IntOp::Add, (false, true), repeats, false),
                3 => compiled(fast, IntOp::Add, (1, 1), repeats),
                4 => compiled(fast, IntOp::Add, (4, 4), repeats),
                _ => compiled(fast, IntOp::Add, (1, 0), repeats),
            };
            let (va, fa) = worker(make(false));
            let (vb, fb) = worker(make(!control));
            let mut vs = [va, vb];
            let fs = [fa, fb];
            let args = match shape {
                0 => [Value::I32(-17), Value::I32(11)],
                1 => [Value::I32(-17), Value::F64(2147483647.)],
                2 => [Value::F64(2147483647.), Value::I32(-17)],
                _ => [Value::F64(-17.), Value::F64(11.)],
            };
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
