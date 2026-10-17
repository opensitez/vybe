use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/math_stack_before.rs"]
mod before;
fn output(name: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&d).unwrap();
    std::fs::write(
        d.join(format!("MATH_STACK_{name}.json")),
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
fn program(fast: bool, kind: usize, repetitions: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("math-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = if kind == 0 { 1 } else { 3 };
    c.param_count = c.arity;
    c.local_count = c.arity as u16;
    for i in 0..repetitions {
        for a in 0..c.arity {
            c.emit_op_u16(Op::LOCAL_GET, a as u16, 0);
        }
        match (kind, fast) {
            (0, false) => before::emit_signbit(&mut c, 0),
            (0, true) => vybe_compiler::primitives::math::emit_signbit(&mut c, 0),
            (_, false) => before::emit_clamp(&mut c, 0),
            (_, true) => vybe_compiler::primitives::math::emit_clamp(&mut c, 0),
        }
        if i + 1 < repetitions {
            c.emit_op(Op::DROP, 0);
        }
    }
    c.emit_op(Op::RETURN, 0);
    root.imports = c.imports.clone();
    vec![root, c]
}
fn worker(cs: Vec<Chunk>) -> (VM, Value) {
    let mut v = VM::new();
    let f = v.run(cs).unwrap();
    (v, f)
}
fn same(a: &Value, b: &Value) {
    match (a, b) {
        (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        _ => assert_eq!(a, b),
    }
}
#[test]
fn bit_patterns_and_clamp_edges() {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap())
        .join("math-native");
    std::fs::create_dir_all(&d).unwrap();
    let mut sizes = vec![];
    for k in 0..2 {
        for fast in [false, true] {
            let cs = program(fast, k, 1);
            let bytes = native(&cs);
            std::fs::write(d.join(format!("{k}-{fast}.wasm")), &bytes).unwrap();
            sizes.push(serde_json::json!({"kind":k,"fast":fast,"body_bytes":cs[1].code.len(),"locals":cs[1].local_count,"native_bytes":bytes.len()}));
        }
    }
    let (mut a, af) = worker(program(false, 0, 1));
    let (mut b, bf) = worker(program(true, 0, 1));
    let mut state = 0x123456789abcdef0u64;
    let mut bits = vec![
        0,
        1,
        0x8000000000000000,
        0x8000000000000001,
        0x7ff0000000000000,
        0xfff0000000000000,
        0x7ff0000000000001,
        0xfff0000000000001,
        0x7ff8000000000001,
        0xfff8000000000001,
        u64::MAX,
    ];
    for i in 0..64 {
        bits.push(1u64 << i);
        bits.push(!(1u64 << i));
    }
    for _ in 0..16000 {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        bits.push(state);
    }
    let mut rows = vec![];
    for n in bits {
        let v = Value::F64(f64::from_bits(n));
        let old = a.invoke(&af, &[v.clone()]).unwrap();
        let new = b.invoke(&bf, &[v]).unwrap();
        assert_eq!(old, new, "{n:016x}");
        assert_eq!(new, Value::Bool(n >> 63 != 0));
        rows.push(format!("{n:016x}"));
    }
    for v in [
        Value::I32(i32::MIN),
        Value::I32(0),
        Value::I32(1),
        Value::Null,
        Value::Undefined,
        Value::Bool(true),
        Value::String("-1".into()),
    ] {
        let old = a.invoke(&af, &[v.clone()]).map_err(|e| e.message);
        let new = b.invoke(&bf, &[v]).map_err(|e| e.message);
        assert_eq!(old, new);
    }
    let (mut a, af) = worker(program(false, 1, 1));
    let (mut b, bf) = worker(program(true, 1, 1));
    let values = [
        f64::NEG_INFINITY,
        -f64::MAX,
        -1.,
        -f64::MIN_POSITIVE,
        -f64::from_bits(1),
        -0.,
        0.,
        f64::from_bits(1),
        f64::MIN_POSITIVE,
        1.,
        f64::MAX,
        f64::INFINITY,
        f64::from_bits(0x7ff0000000000001),
        f64::from_bits(0xfff8000000000001),
    ];
    let mut clamps = vec![];
    for x in values {
        for min in values {
            for max in values {
                let args = [Value::F64(x), Value::F64(min), Value::F64(max)];
                same(
                    &a.invoke(&af, &args).unwrap(),
                    &b.invoke(&bf, &args).unwrap(),
                );
                clamps.push([
                    format!("{:016x}", x.to_bits()),
                    format!("{:016x}", min.to_bits()),
                    format!("{:016x}", max.to_bits()),
                ]);
            }
        }
    }
    std::fs::write(
        d.join("queries.json"),
        serde_json::to_vec(&serde_json::json!({"signbits":rows,"clamps":clamps})).unwrap(),
    )
    .unwrap();
    output(
        "GATES",
        &serde_json::json!({"sign_bits":rows.len(),"coercions":7,"clamps":clamps.len(),"sizes":sizes}),
    );
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for repetitions in [1, 32] {
        for (op, label, args, control) in [
            (0, "sign-positive", vec![Value::F64(123.5)], false),
            (0, "sign-negative", vec![Value::F64(-123.5)], false),
            (
                0,
                "sign-nan",
                vec![Value::F64(f64::from_bits(0xfff8000000000001))],
                false,
            ),
            (0, "sign-control", vec![Value::F64(-123.5)], true),
            (
                1,
                "clamp-inside",
                vec![Value::F64(0.5), Value::F64(0.), Value::F64(1.)],
                false,
            ),
            (
                1,
                "clamp-outside",
                vec![Value::F64(-5.), Value::F64(0.), Value::F64(1.)],
                false,
            ),
            (
                1,
                "clamp-nan",
                vec![Value::F64(f64::NAN), Value::F64(0.), Value::F64(1.)],
                false,
            ),
            (
                1,
                "clamp-control",
                vec![Value::F64(0.5), Value::F64(0.), Value::F64(1.)],
                true,
            ),
        ] {
            let old = program(false, op, repetitions);
            let new = program(!control, op, repetitions);
            let sizes = serde_json::json!({"old_guest":old[1].code.len(),"new_guest":new[1].code.len(),"old_native":native(&old).len(),"new_native":native(&new).len(),"old_locals":old[1].local_count,"new_locals":new[1].local_count});
            let (va, fa) = worker(old);
            let (vb, fb) = worker(new);
            let mut vs = [va, vb];
            let fs = [fa, fb];
            let expected = vs[0].invoke(&fs[0], &args).unwrap();
            same(&vs[1].invoke(&fs[1], &args).unwrap(), &expected);
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
            rows.push(serde_json::json!({"label":label,"op":op,"repetitions":repetitions,"sizes":sizes,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
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
