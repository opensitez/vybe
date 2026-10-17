use vybe_ast::FloatLane;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/nextafter_copysign_before.rs"]
mod before;
#[path="fixtures/nextafter_copysign_candidate.rs"]
mod candidate;
fn output(name: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&d).unwrap();
    std::fs::write(
        d.join(format!("NEXT_COPYSIGN_{name}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
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
fn program(fast: bool, lane: FloatLane, reps: usize) -> Vec<Chunk> {
    program_kind(fast, lane, reps, 0)
}
fn program_kind(fast: bool, lane: FloatLane, reps: usize, kind: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("next-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    for i in 0..reps {
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        if kind == 0 { c.emit_op_u16(Op::LOCAL_GET, 1, 0); }
        match (fast, kind) {
            (false, 0) => before::emit_next_after(&mut c, lane, 0),
            (true, 0) => vybe_compiler::primitives::math::emit_next_after(&mut c, lane, 0),
            (false, 1 | 2) => before::emit_next_toward(&mut c, kind == 1, lane, 0),
            (true, 1 | 2) => vybe_compiler::primitives::math::emit_next_toward(&mut c, kind == 1, lane, 0),
            (false, 3) => before::emit_ulp(&mut c, lane, 0),
            (true, 3) => vybe_compiler::primitives::math::emit_ulp(&mut c, lane, 0),
            _ => unreachable!(),
        }
        if i + 1 < reps {
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
fn exact_results_including_nan_payloads() {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap())
        .join("next-native");
    std::fs::create_dir_all(&d).unwrap();
    let mut pairs = vec![];
    let v = [
        0.,
        -0.,
        1.,
        -1.,
        2.,
        -2.,
        f64::MIN_POSITIVE,
        -f64::MIN_POSITIVE,
        f64::from_bits(1),
        -f64::from_bits(1),
        f64::MAX,
        -f64::MAX,
        f64::INFINITY,
        f64::NEG_INFINITY,
        f64::from_bits(0x7ff0000000000001),
        f64::from_bits(0xfff8000000000042),
    ];
    for x in v {
        for y in v {
            pairs.push([x.to_bits(), y.to_bits()]);
        }
    }
    let mut state = 0x123456789abcdef0u64;
    for _ in 0..12000 {
        let mut pair = [0u64; 2];
        for x in &mut pair {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            *x = state;
        }
        pairs.push(pair);
    }
    let mut sizes = vec![];
    let mut checks = 0;
    for (k, lane) in [FloatLane::F32, FloatLane::F64].into_iter().enumerate() {
        let (mut a, af) = worker(program(false, lane, 1));
        let (mut b, bf) = worker(program(true, lane, 1));
        for p in &pairs {
            let args = [
                Value::F64(f64::from_bits(p[0])),
                Value::F64(f64::from_bits(p[1])),
            ];
            let x = a.invoke(&af, &args).unwrap().as_f64();
            let y = b.invoke(&bf, &args).unwrap().as_f64();
            assert_eq!(x.to_bits(), y.to_bits(), "{lane:?} {p:?}");
            checks += 1;
        }
        for fast in [false, true] {
            let cs = program(fast, lane, 1);
            let bytes = native(&cs);
            std::fs::write(d.join(format!("{k}-{fast}.wasm")), &bytes).unwrap();
            sizes.push(serde_json::json!({"lane":k,"fast":fast,"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
        }
    }
    let rows: Vec<_> = pairs
        .iter()
        .map(|p| [format!("{:016x}", p[0]), format!("{:016x}", p[1])])
        .collect();
    std::fs::write(d.join("queries.json"), serde_json::to_vec(&rows).unwrap()).unwrap();
    output(
        "GATES",
        &serde_json::json!({"comparisons":checks,"pairs":pairs.len(),"sizes":sizes}),
    );
}
#[test]
fn unary_family_and_carriers() {
    let dir = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap()).join("unary-native");
    std::fs::create_dir_all(&dir).unwrap();
    let mut bits = vec![0,1,0x000fffffffffffff,0x0010000000000000,0x3ff0000000000000,0x7fefffffffffffff,0x7ff0000000000000,0x7ff0000000000001,0x7ff8000000000042];
    bits.extend(bits.clone().into_iter().map(|x|x|0x8000000000000000));
    let mut state=0x123456789abcdef0u64;
    for _ in 0..12000 { state^=state<<13;state^=state>>7;state^=state<<17;bits.push(state); }
    // f32 rounding, underflow and overflow boundaries, including half-way values.
    for word in [0u32,1,0x007fffff,0x00800000,0x3f800000,0x7f7fffff,0x7f800000,0x7f800001] {
        for sign in [0,0x80000000] {
            let x=f32::from_bits(word|sign) as f64;
            bits.push(x.to_bits());
            for delta in [-1i64,1] { bits.push(x.to_bits().wrapping_add_signed(delta)); }
        }
    }
    let carriers = vec![Value::I32(1),Value::I32(-1),Value::I64(i64::MAX),Value::Bool(true),Value::Bool(false),Value::Null,Value::Undefined,Value::String("1".into()),Value::String("bad".into())];
    let mut checks=0;let mut carrier_checks=0;let mut sizes=vec![];
    for kind in 0..4 {
        for (lane_id,lane) in [FloatLane::F32,FloatLane::F64].into_iter().enumerate() {
            let old=program_kind(false,lane,1,kind); let new=program_kind(true,lane,1,kind);
            for (fast,cs) in [(false,&old),(true,&new)] {
                let bytes=native(cs);
                std::fs::write(dir.join(format!("{kind}-{lane_id}-{fast}.wasm")),&bytes).unwrap();
                sizes.push(serde_json::json!({"kind":kind,"lane":lane_id,"fast":fast,"guest":cs[1].code.len(),"native":bytes.len()}));
            }
            let (mut a,fa)=worker(old);let(mut b,fb)=worker(new);
            for &word in &bits {
                let args=[Value::F64(f64::from_bits(word)),Value::F64(2.0)];
                let x=a.invoke(&fa,&args).unwrap().as_f64();let y=b.invoke(&fb,&args).unwrap().as_f64();
                assert_eq!(x.to_bits(),y.to_bits(),"kind={kind} lane={lane:?} bits={word:016x}");checks+=1;
            }
            for value in &carriers {
                for args in [[value.clone(),Value::F64(2.0)],[Value::F64(1.0),value.clone()]] {
                    let old=a.invoke(&fa,&args).map(|x|x.as_f64().to_bits()).map_err(|e|e.message);
                    let new=b.invoke(&fb,&args).map(|x|x.as_f64().to_bits()).map_err(|e|e.message);
                    assert_eq!(old,new,"kind={kind} lane={lane:?} args={args:?}");carrier_checks+=1;
                }
            }
        }
    }
    std::fs::write(dir.join("queries.json"),serde_json::to_vec(&bits.iter().map(|x|format!("{x:016x}")).collect::<Vec<_>>()).unwrap()).unwrap();
    output("FAMILY",&serde_json::json!({"checks":checks,"carrier_checks":carrier_checks,"sizes":sizes}));
}

#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for repetitions in [1, 16] {
        for lane in [FloatLane::F32, FloatLane::F64] {
            for (label, x, y, control) in [
                ("up-positive", 1., 2., false),
                ("up-negative", -2., -1., false),
                ("up-zero", 0., 1., false),
                ("down", 2., 1., false),
                ("equal", 1., 1., false),
                ("nan", f64::from_bits(0xfff8000000000042), 1., false),
                ("control", 1., 2., true),
            ] {
                let args = [Value::F64(x), Value::F64(y)];
                let old = program(false, lane, repetitions);
                let new = program(!control, lane, repetitions);
                let sizes = serde_json::json!({"old_guest":old[1].code.len(),"new_guest":new[1].code.len(),"old_native":native(&old).len(),"new_native":native(&new).len(),"old_locals":old[1].local_count,"new_locals":new[1].local_count});
                let (va, fa) = worker(old);
                let (vb, fb) = worker(new);
                let mut vs = [va, vb];
                let fs = [fa, fb];
                let expected = vs[0].invoke(&fs[0], &args).unwrap();
                assert_eq!(
                    vs[1].invoke(&fs[1], &args).unwrap().as_f64().to_bits(),
                    expected.as_f64().to_bits()
                );
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
                let mut ratios: Vec<_> = wall[0].iter().zip(&wall[1]).map(|(a, b)| a / b).collect();
                ratios.sort_by(f64::total_cmp);
                let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(a, b)| a / b).collect();
                cr.sort_by(f64::total_cmp);
                rows.push(serde_json::json!({"label":label,"lane":format!("{lane:?}"),"repetitions":repetitions,"sizes":sizes,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
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
