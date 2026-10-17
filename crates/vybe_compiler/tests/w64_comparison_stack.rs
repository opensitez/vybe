use vybe_ast::{BitLane, IntOp};
use vybe_compiler::primitives::bits;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/w64_comparison_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
const CMP: [IntOp; 10] = [
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
        p.join(format!("W64_COMPARE_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn program(fast: bool, op: IntOp, repeats: usize, prefix: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("wide-compare-owner");
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
            bits::emit_integer_binary(&mut c, BitLane::W64, op, 0);
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
fn worker(cs: Vec<Chunk>) -> (VM, Value) {
    let mut v = VM::new();
    let f = v.run(cs).unwrap();
    (v, f)
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
fn word(x: f64) -> u64 {
    if x >= 9223372036854775808.0 {
        x as u64
    } else {
        (x as i64) as u64
    }
}
fn oracle(op: IntOp, a: f64, b: f64) -> bool {
    let (a, b) = (word(a), word(b));
    match op {
        IntOp::Eq => a == b,
        IntOp::Ne => a != b,
        IntOp::LtS => (a as i64) < (b as i64),
        IntOp::LtU => a < b,
        IntOp::LeS => (a as i64) <= (b as i64),
        IntOp::LeU => a <= b,
        IntOp::GtS => (a as i64) > (b as i64),
        IntOp::GtU => a > b,
        IntOp::GeS => (a as i64) >= (b as i64),
        IntOp::GeU => a >= b,
        _ => unreachable!(),
    }
}
#[test]
fn other_lanes_and_arithmetic_keep_the_exact_original_emission() {
    let arithmetic = [
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
    ];
    for lane in [BitLane::W32, BitLane::W64] {
        for op in arithmetic.into_iter().chain(
            if lane == BitLane::W32 {
                CMP.as_slice()
            } else {
                &[]
            }
            .iter()
            .copied(),
        ) {
            for normalized in [(false, false), (false, true), (true, false), (true, true)] {
                for locals in [2, 11] {
                    let mut a = Chunk::new("original");
                    a.local_count = locals;
                    let mut b = a.clone();
                    before::emit_integer_binary_with_normalized_operands(
                        &mut a, lane, op, normalized, 0,
                    );
                    bits::emit_integer_binary_with_normalized_operands(
                        &mut b, lane, op, normalized, 0,
                    );
                    assert_eq!(a.code, b.code, "{lane:?} {op:?} {normalized:?}");
                    assert_eq!(a.local_count, b.local_count);
                    assert_eq!(a.i64_scratch_slot, b.i64_scratch_slot);
                    assert_eq!(a.imports, b.imports);
                }
            }
        }
    }
}
#[test]
fn signed_unsigned_boundaries_random_values_and_traps() {
    let points = [
        -9223372036854775808.0,
        -9223372036854774784.0,
        -4503599627370496.0,
        -4294967296.0,
        -2147483648.0,
        -1.5,
        -1.0,
        -0.5,
        -0.0,
        0.0,
        0.5,
        1.0,
        1.5,
        2147483647.0,
        2147483648.0,
        4294967295.0,
        4294967296.0,
        4503599627370496.0,
        9007199254740991.0,
        9223372036854774784.0,
        9223372036854775808.0,
        9223372036854777856.0,
        18446744073709549568.0,
    ];
    let mut pairs = vec![];
    for a in points {
        for b in points {
            pairs.push((a, b));
        }
    }
    let mut seed = 0x132456789abcdef0u64;
    for _ in 0..4096 {
        seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1);
        let a = if seed & 1 == 0 {
            seed as f64
        } else {
            (seed as i64) as f64
        };
        seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1);
        let b = if seed & 1 == 0 {
            seed as f64
        } else {
            (seed as i64) as f64
        };
        if a < 18446744073709551616.0 && b < 18446744073709551616.0 {
            pairs.push((a, b));
        }
    }
    let mut checks = 0;
    let mut sizes = vec![];
    let mut corpus = vec![];
    for (index, op) in CMP.into_iter().enumerate() {
        let old = program(false, op, 1, false);
        let new = program(true, op, 1, false);
        let bytes = [native(&old), native(&new)];
        assert!(new[1].code.len() < old[1].code.len());
        assert!(new[1].local_count < old[1].local_count);
        native_export::export_chunks(&old, &format!("w64-compare-{index}-before"));
        native_export::export_chunks(&new, &format!("w64-compare-{index}-after"));
        sizes.push(serde_json::json!({"op":format!("{op:?}"),"old_body":old[1].code.len(),"new_body":new[1].code.len(),"old_locals":old[1].local_count,"new_locals":new[1].local_count,"old_native":bytes[0].len(),"new_native":bytes[1].len()}));
        let (va, fa) = worker(old);
        let (vb, fb) = worker(new);
        let mut vs = [va, vb];
        let fs = [fa, fb];
        for &(a, b) in &pairs {
            let expected = oracle(op, a, b);
            for side in 0..2 {
                let out = vs[side]
                    .invoke(&fs[side], &[Value::F64(a), Value::F64(b)])
                    .unwrap();
                assert_eq!(out, Value::I32(i32::from(expected)), "{op:?} {a} {b}");
                checks += 1;
            }
        }
        for a in &points {
            for b in &points {
                corpus
                    .push(serde_json::json!({"op":index,"a":a,"b":b,"expected":oracle(op,*a,*b)}));
            }
        }
        let bad = [
            Value::F64(f64::NAN),
            Value::F64(f64::INFINITY),
            Value::F64(f64::NEG_INFINITY),
            Value::F64(-9223372036854777856.0),
            Value::F64(18446744073709551616.0),
            Value::Null,
            Value::Undefined,
            Value::Bool(true),
            Value::String("7".into()),
            Value::I32(-7),
            Value::F32(3.5),
        ];
        for a in &bad {
            for b in &bad {
                let mut results = vec![];
                for fast in [false, true] {
                    let (mut v, f) = worker(program(fast, op, 1, false));
                    results.push(v.invoke(&f, &[a.clone(), b.clone()]).map_err(|e| {
                        let trap = e.is_trap();
                        (e.message, trap)
                    }));
                    checks += 1;
                }
                assert_eq!(results[0], results[1], "{op:?} {a:?} {b:?}");
            }
        }
        for fast in [false, true] {
            let cs = program(fast, op, 1, true);
            native(&cs);
            let (mut v, f) = worker(cs);
            assert_eq!(
                v.invoke(&f, &[Value::F64(9223372036854775808.0), Value::F64(-1.)])
                    .unwrap(),
                Value::I32(23)
            );
            checks += 1;
        }
    }
    output(
        "GATES",
        &serde_json::json!({"checks":checks,"valid_pairs_per_op":pairs.len(),"sizes":sizes}),
    );
    output("CORPUS", &serde_json::json!({"queries":corpus}));
}

#[test]
fn actual_compiler_integer_expressions_use_the_wide_path() {
    use vybe_ast::{BinOp, ExprKind, Expression, Lang, Literal, Module, Statement, StmtKind};
    let mut checks = 0;
    for (index, op) in CMP.into_iter().enumerate() {
        for (case, (a, b)) in [
            (-1.5, 0.0),
            (-9223372036854775808.0, 1.0),
            (9223372036854775808.0, -1.0),
            (18446744073709549568.0, 9223372036854775808.0),
            (7.0, 7.0),
        ]
        .into_iter()
        .enumerate()
        {
            let p = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
            let expr = Expression::new(ExprKind::Binary {
                op: BinOp::Integer(op, BitLane::W64),
                left: Box::new(Expression::new(ExprKind::Lit(Literal::Float(a)))),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Float(b)))),
            });
            let module = Module {
                name: "wide-compiler-proof".into(),
                language: Lang::Unknown,
                body: vec![Statement::new(StmtKind::Return(Some(expr)))],
                imports: vec![],
                canon: Default::default(),
                directives: Default::default(),
            };
            let mut cs = vybe_compiler::Compiler::with_profile(p)
                .compile(&module)
                .unwrap();
            native(&cs);
            let mut v = VM::new();
            assert_eq!(
                v.run(cs.clone()).unwrap().as_i32(),
                i32::from(oracle(op, a, b))
            );
            cs[0].name = "entry".into();
            native_export::export_chunks(&cs, &format!("w64-compiled-{index}-{case}"));
            checks += 1;
        }
    }
    output("COMPILER", &serde_json::json!({"checks":checks}));
}

#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, op, a, b, control) in [
        ("signed-low", IntOp::LtS, 7.0, 11.0, false),
        ("unsigned-low", IntOp::LtU, 7.0, 11.0, false),
        (
            "signed-high",
            IntOp::LtS,
            9223372036854775808.0,
            18446744073709549568.0,
            false,
        ),
        (
            "unsigned-high",
            IntOp::LtU,
            9223372036854775808.0,
            18446744073709549568.0,
            false,
        ),
        (
            "signed-mixed",
            IntOp::LtS,
            -17.0,
            9223372036854775808.0,
            false,
        ),
        (
            "equality",
            IntOp::Eq,
            9223372036854775808.0,
            9223372036854775808.0,
            false,
        ),
        ("control", IntOp::LtS, 7.0, 11.0, true),
    ] {
        for repeats in [1, 16] {
            let (va, fa) = worker(program(false, op, repeats, false));
            let (vb, fb) = worker(program(!control, op, repeats, false));
            let mut vs = [va, vb];
            let fs = [fa, fb];
            let args = [Value::F64(a), Value::F64(b)];
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
