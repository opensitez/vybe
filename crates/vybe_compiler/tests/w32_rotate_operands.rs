use vybe_ast::{BinOp, BitLane, ExprKind, Expression, Lang, Module, Statement, StmtKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/rotate_before.rs"]
mod before;
fn output(name: &str, v: &serde_json::Value) {
    let d = std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap();
    std::fs::write(
        std::path::Path::new(&d).join(format!("ROTATE_{name}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn native(chunks: &[Chunk]) -> Vec<u8> {
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: false,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    bytes
}
fn compiled(lane: BitLane, left: bool, a: Expression, b: Expression) -> Vec<Chunk> {
    let p = vybe_compiler::profile::parse_profile("[compiler]\n").unwrap();
    let m = Module {
        name: "rotate-proof".into(),
        language: Lang::Unknown,
        body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
            ExprKind::Binary {
                op: if left {
                    BinOp::RotL(lane)
                } else {
                    BinOp::RotR(lane)
                },
                left: Box::new(a),
                right: Box::new(b),
            },
        ))))],
        imports: vec![],
        canon: Default::default(),
        directives: Default::default(),
    };
    vybe_compiler::Compiler::with_profile(p)
        .compile(&m)
        .unwrap()
}
fn run(chunks: Vec<Chunk>) -> Value {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap()
}
fn helper(
    fast: bool,
    lane: BitLane,
    left: bool,
    proofs: (bool, bool),
    a: f64,
    b: f64,
    repetitions: usize,
) -> Vec<Chunk> {
    let mut root = Chunk::new("rotate-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("rotate");
    for i in 0..repetitions {
        for (value, proof) in [(a, proofs.0), (b, proofs.1)] {
            if proof {
                c.emit_i32_const(value as i32, 0);
            } else {
                c.emit_f64_const(value, 0);
            }
        }
        if fast {
            vybe_compiler::primitives::bits::emit_rotate_with_normalized_operands(
                &mut c, lane, left, proofs, 0,
            );
        } else {
            before::emit_rotate_with_normalized_operands(&mut c, lane, left, proofs, 0);
        }
        if i + 1 < repetitions {
            c.emit_op(Op::DROP, 0);
        }
    }
    c.emit_op(Op::RETURN, 0);
    vec![root, c]
}
fn worker(chunks: Vec<Chunk>) -> (VM, Value) {
    let mut v = VM::new();
    let f = v.run(chunks).unwrap();
    (v, f)
}
#[test]
fn literal_words_use_native_rotate_without_spills() {
    let mut checks = 0;
    let mut rows = vec![];
    let dir = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap())
        .join("rotate-native");
    std::fs::create_dir_all(&dir).unwrap();
    for a in [
        i32::MIN as i64,
        -1,
        0,
        1,
        i32::MAX as i64,
        0x8000_0000,
        0xffff_ffff,
    ] {
        for b in [
            i32::MIN as i64,
            -65,
            -33,
            -32,
            -1,
            0,
            1,
            31,
            32,
            33,
            65,
            i32::MAX as i64,
            0x8000_0000,
            0xffff_ffff,
        ] {
            for left in [false, true] {
                let c = compiled(BitLane::W32, left, Expression::int(a), Expression::int(b));
                let dump = vybe_runtime::debug::disassemble(&c[0]);
                assert!(!dump.contains("trunc_f64"), "{dump}");
                assert!(!dump.contains("wrap_i64"), "{dump}");
                assert!(!dump.contains("local.set"), "{dump}");
                native(&c);
                let expected = if left {
                    (a as u32).rotate_left(b as u32)
                } else {
                    (a as u32).rotate_right(b as u32)
                };
                let name = format!("case-{checks}.wasm");
                std::fs::write(dir.join(&name), native(&c)).unwrap();
                rows.push(serde_json::json!({"file":name,"expected":expected as i32}));
                assert_eq!(run(c).as_i32(), expected as i32);
                checks += 1;
            }
        }
    }
    output(
        "LITERAL_GATES",
        &serde_json::json!({"checks":checks,"native_validation":true}),
    );
    std::fs::write(dir.join("cases.json"), serde_json::to_vec(&rows).unwrap()).unwrap();
}
#[test]
fn mixed_operands_and_wide_fallback_match_frozen_lowering() {
    let mut checks = 0;
    let mut sizes = vec![];
    for lane in [BitLane::W32, BitLane::W64] {
        for left in [false, true] {
            for proofs in [(false, false), (false, true), (true, false), (true, true)] {
                if lane == BitLane::W64 && proofs != (false, false) {
                    continue;
                }
                for a in [-2147483648., -1., 0., 1., 2147483647.] {
                    for b in [-65., -1., 0., 1., 32., 65.] {
                        let old = helper(false, lane, left, proofs, a, b, 1);
                        let new = helper(true, lane, left, proofs, a, b, 1);
                        let old_native = native(&old);
                        let new_native = native(&new);
                        let (mut v, f) = worker(old.clone());
                        let expected = v.invoke(&f, &[]).unwrap();
                        let (mut v, f) = worker(new.clone());
                        assert_eq!(v.invoke(&f, &[]).unwrap(), expected);
                        if lane == BitLane::W64 {
                            assert_eq!(old[1].code, new[1].code);
                        }
                        if a == 1. && b == 1. {
                            sizes.push(serde_json::json!({"lane":format!("{lane:?}"),"left":left,"proofs":proofs,"old_bytes":old[1].code.len(),"new_bytes":new[1].code.len(),"old_native":old_native.len(),"new_native":new_native.len(),"old_locals":old[1].local_count,"new_locals":new[1].local_count}));
                        }
                        checks += 1;
                    }
                }
            }
        }
    }
    for value in [i32::MIN as f64 - 1., u32::MAX as f64 + 1., 4294967295.75] {
        let c = compiled(
            BitLane::W32,
            true,
            Expression::float(value),
            Expression::int(1),
        );
        let dump = vybe_runtime::debug::disassemble(&c[0]);
        assert!(dump.contains("trunc_f64"), "{dump}");
        native(&c);
        let (mut v, f) = worker(helper(
            false,
            BitLane::W32,
            true,
            (false, false),
            value,
            1.,
            1,
        ));
        let expected = v.invoke(&f, &[]).unwrap();
        assert_eq!(run(c), expected);
        checks += 1;
    }
    output(
        "HELPER_GATES",
        &serde_json::json!({"checks":checks,"sizes":sizes}),
    );
}

#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for repetitions in [1, 32] {
        for (label, lane, old_proofs, new_proofs, control) in [
            (
                "control",
                BitLane::W32,
                (false, false),
                (false, false),
                true,
            ),
            (
                "unproven",
                BitLane::W32,
                (false, false),
                (false, false),
                false,
            ),
            (
                "left-word",
                BitLane::W32,
                (true, false),
                (true, false),
                false,
            ),
            (
                "count-word",
                BitLane::W32,
                (false, true),
                (false, true),
                false,
            ),
            (
                "both-words",
                BitLane::W32,
                (true, true),
                (true, true),
                false,
            ),
            (
                "literal-operands",
                BitLane::W32,
                (false, false),
                (true, true),
                false,
            ),
            (
                "w64-control",
                BitLane::W64,
                (false, false),
                (false, false),
                false,
            ),
        ] {
            let old = helper(false, lane, true, old_proofs, 123456789., 7., repetitions);
            let new = helper(
                !control,
                lane,
                true,
                new_proofs,
                123456789.,
                7.,
                repetitions,
            );
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
            rows.push(serde_json::json!({"label":label,"repetitions":repetitions,"sizes":sizes,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
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
fn nested_rotates_report_their_actual_i32_result() {
    let inner = Expression::new(ExprKind::Binary {
        op: BinOp::RotR(BitLane::W32),
        left: Box::new(Expression::int(0xdead_beef)),
        right: Box::new(Expression::int(7)),
    });
    for (left, b) in [(true, 33), (false, -1)] {
        let c = compiled(BitLane::W32, left, inner.clone(), Expression::int(b));
        let dump = vybe_runtime::debug::disassemble(&c[0]);
        assert!(!dump.contains("trunc_f64"), "{dump}");
        assert!(!dump.contains("local.set"), "{dump}");
        native(&c);
        let x = 0xdead_beefu32.rotate_right(7);
        let expected = if left {
            x.rotate_left(b as u32)
        } else {
            x.rotate_right(b as u32)
        };
        assert_eq!(run(c).as_i32(), expected as i32);
    }
}
