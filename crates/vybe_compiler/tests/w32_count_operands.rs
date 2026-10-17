#[path = "fixtures/count_before.rs"]
mod before;
use vybe_ast::{BinOp, BitLane, ExprKind, Expression, Lang, Module, Statement, StmtKind, UnaryOp};
use vybe_runtime::{Chunk, Op, VM, Value};
const OPS: [(Op, Op); 3] = [
    (Op::I32_POPCNT, Op::I64_POPCNT),
    (Op::I32_CLZ, Op::I64_CLZ),
    (Op::I32_CTZ, Op::I64_CTZ),
];
fn output(name: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&d).unwrap();
    std::fs::write(
        d.join(format!("COUNT_{name}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn unary(op: usize, lane: BitLane, inner: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: match op {
            0 => UnaryOp::PopCount(lane),
            1 => UnaryOp::LeadingZeros(lane),
            _ => UnaryOp::TrailingZeros(lane),
        },
        expr: Box::new(inner),
    })
}
fn compiled(e: Expression) -> Vec<Chunk> {
    let p = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
    let m = Module {
        name: "count-proof".into(),
        language: Lang::Unknown,
        body: vec![Statement::new(StmtKind::Return(Some(e)))],
        imports: vec![],
        canon: Default::default(),
        directives: Default::default(),
    };
    vybe_compiler::Compiler::with_profile(p)
        .compile(&m)
        .unwrap()
}
fn native(cs: &[Chunk]) -> Vec<u8> {
    let b = vybe_platform_wasm::write_wasm_with_options(
        cs,
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
fn run(cs: Vec<Chunk>) -> Value {
    let mut v = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut v);
    v.run(cs).unwrap()
}
fn expected(op: usize, n: u32) -> i32 {
    match op {
        0 => n.count_ones() as i32,
        1 => n.leading_zeros() as i32,
        _ => n.trailing_zeros() as i32,
    }
}
fn helper(
    fast: bool,
    op: usize,
    lane: BitLane,
    normalized: bool,
    value: f64,
    repetitions: usize,
) -> Vec<Chunk> {
    let mut root = Chunk::new("count-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("count");
    for i in 0..repetitions {
        if fast && normalized && lane == BitLane::W32 {
            c.emit_i32_const(value as i64 as i32, 0);
        } else {
            c.emit_f64_const(value, 0);
        }
        if fast {
            vybe_compiler::primitives::bits::emit_count_with_normalized_operand(
                &mut c, lane, OPS[op].0, OPS[op].1, normalized, 0,
            );
        } else {
            before::emit_count(&mut c, lane, OPS[op].0, OPS[op].1, 0);
        }
        if i + 1 < repetitions {
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
#[test]
fn literal_and_nested_count_cases() {
    let mut values: Vec<i64> = (-128..=255).collect();
    values.extend([
        i32::MIN as i64,
        i32::MAX as i64,
        0x8000_0000,
        0xffff_ffff,
        0x55555555,
        0xaaaaaaaa,
    ]);
    for b in 0..32 {
        values.push(1i64 << b);
        values.push((u32::MAX ^ (1u32 << b)) as i64);
    }
    let dir = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap())
        .join("count-native");
    std::fs::create_dir_all(&dir).unwrap();
    let mut rows = vec![];
    let mut checks = 0;
    for op in 0..3 {
        for &value in &values {
            let cs = compiled(unary(op, BitLane::W32, Expression::int(value)));
            let dump = vybe_runtime::debug::disassemble(&cs[0]);
            assert!(!dump.contains("trunc_f64"), "{dump}");
            assert!(!dump.contains("wrap_i64"), "{dump}");
            let bytes = native(&cs);
            let want = expected(op, value as u32);
            assert_eq!(run(cs).as_i32(), want);
            let file = format!("case-{checks}.wasm");
            std::fs::write(dir.join(&file), bytes).unwrap();
            rows.push(serde_json::json!({"file":file,"expected":want}));
            checks += 1;
        }
        for inner in [
            unary(0, BitLane::W32, Expression::int(0xffff_ffff)),
            Expression::new(ExprKind::Binary {
                op: BinOp::RotL(BitLane::W32),
                left: Box::new(Expression::int(7)),
                right: Box::new(Expression::int(1)),
            }),
            Expression::new(ExprKind::Binary {
                op: BinOp::Integer(vybe_ast::IntOp::Add, BitLane::W32),
                left: Box::new(Expression::int(5)),
                right: Box::new(Expression::int(6)),
            }),
        ] {
            let input = run(compiled(inner.clone())).as_i32();
            let cs = compiled(unary(op, BitLane::W32, inner));
            let dump = vybe_runtime::debug::disassemble(&cs[0]);
            assert!(!dump.contains("trunc_f64"), "{dump}");
            native(&cs);
            assert_eq!(run(cs).as_i32(), expected(op, input as u32));
            checks += 1;
        }
    }
    std::fs::write(dir.join("cases.json"), serde_json::to_vec(&rows).unwrap()).unwrap();
    output(
        "GATES",
        &serde_json::json!({"checks":checks,"native_cases":rows.len()}),
    );
}
#[test]
fn float_and_w64_paths_retain_conversion() {
    let mut checks = 0;
    for op in 0..3 {
        for value in [
            -2147483649.,
            4294967296.,
            4294967295.75,
            -1.75,
            0.,
            1.,
            9223372036854774784.,
        ] {
            for lane in [BitLane::W32, BitLane::W64] {
                let cs = compiled(unary(op, lane, Expression::float(value)));
                let dump = vybe_runtime::debug::disassemble(&cs[0]);
                assert!(dump.contains("trunc_f64"), "{dump}");
                native(&cs);
                let (mut v, f) = worker(helper(false, op, lane, false, value, 1));
                assert_eq!(run(cs), v.invoke(&f, &[]).unwrap());
                let a = helper(false, op, lane, false, value, 1);
                let b = helper(true, op, lane, false, value, 1);
                assert_eq!(a[1].code, b[1].code);
                checks += 1;
            }
        }
    }
    output("FALLBACK", &serde_json::json!({"checks":checks}));
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for repetitions in [1, 32] {
        for op in 0..3 {
            for (label, lane, normalized, control) in [
                ("normalized", BitLane::W32, true, false),
                ("unproven-control", BitLane::W32, false, false),
                ("identical-control", BitLane::W32, false, true),
                ("w64-control", BitLane::W64, false, false),
            ] {
                let old = helper(false, op, lane, normalized, 123456789., repetitions);
                let new = helper(!control, op, lane, normalized, 123456789., repetitions);
                let sizes = serde_json::json!({"old_guest":old[1].code.len(),"new_guest":new[1].code.len(),"old_native":native(&old).len(),"new_native":native(&new).len(),"old_locals":old[1].local_count,"new_locals":new[1].local_count});
                let (va, fa) = worker(old);
                let (vb, fb) = worker(new);
                let mut vs = [va, vb];
                let fs = [fa, fb];
                let expected = vs[0].invoke(&fs[0], &[]).unwrap();
                assert_eq!(vs[1].invoke(&fs[1], &[]).unwrap(), expected);
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
                            std::hint::black_box(vs[side].invoke(&fs[side], &[]).unwrap());
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
                let mut ratios: Vec<_> = wall[0].iter().zip(&wall[1]).map(|(a, b)| a / b).collect();
                ratios.sort_by(f64::total_cmp);
                let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(a, b)| a / b).collect();
                cr.sort_by(f64::total_cmp);
                rows.push(serde_json::json!({"label":label,"op":op,"repetitions":repetitions,"sizes":sizes,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
            }
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
