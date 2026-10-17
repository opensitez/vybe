use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/is_inf_before.rs"]
mod before;
fn output(name: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&d).unwrap();
    std::fs::write(
        d.join(format!("IS_INF_{name}.json")),
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
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    for i in 0..repetitions {
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        if fast {
            vybe_compiler::primitives::math::emit_is_inf(&mut c, 0);
        } else {
            before::emit_before(&mut c, 0);
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
    vybe_compiler::primitives::platforms::register_platforms_all(&mut v);
    let f = v.run(cs).unwrap();
    (v, f)
}
fn same(a: &Value, b: &Value) {
    match (a, b) {
        (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        _ => assert_eq!(a, b),
    }
}

#[path = "fixtures/primitive_native_export.rs"]
mod native_export;

fn pairs() -> Vec<(u64, u64)> {
    let edge = [
        0,
        1,
        0x000fffffffffffff,
        0x0010000000000000,
        0x3ff0000000000000,
        0x7fefffffffffffff,
        0x7ff0000000000000,
        0x7ff0000000000001,
        0x7ff8000000000001,
    ];
    let mut values = vec![];
    for n in edge {
        values.push(n);
        values.push(n | 0x8000000000000000);
    }
    let mut out = vec![];
    for &a in &values {
        for &b in &values {
            out.push((a, b));
        }
    }
    let mut n = 0x123456789abcdef0u64;
    for _ in 0..4096 {
        n ^= n << 13;
        n ^= n >> 7;
        n ^= n << 17;
        let a = n;
        n ^= n << 13;
        n ^= n >> 7;
        n ^= n << 17;
        out.push((a, n));
        for x in [
            0x7ff0000000000000,
            0xfff0000000000000,
            0,
            0x8000000000000000,
            0x7ff0000000000001,
            0xfff8000000000001,
        ] {
            out.push((x, n));
        }
    }
    out
}
fn oracle(x: u64, s: u64) -> bool {
    let mask = 0x7fffffffffffffffu64;
    let magnitude = x & mask;
    let sign_magnitude = s & mask;
    magnitude == 0x7ff0000000000000
        && sign_magnitude <= 0x7ff0000000000000
        && (sign_magnitude == 0 || (x >> 63) == (s >> 63))
}
#[test]
fn exact_infinity_and_sign_domain() {
    let cases = pairs();
    let mut sizes = vec![];
    for fast in [false, true] {
        let cs = program(fast, 0, 1);
        let bytes = native(&cs);
        sizes.push(serde_json::json!({"fast":fast,"guest_bytes":cs[1].code.len(),"locals":cs[1].local_count,"native_bytes":bytes.len()}));
        native_export::export_chunks(
            &cs,
            if fast {
                "is-inf-after"
            } else {
                "is-inf-before"
            },
        );
        let (mut v, f) = worker(cs);
        for &(a, b) in &cases {
            let got = v
                .invoke(
                    &f,
                    &[Value::F64(f64::from_bits(a)), Value::F64(f64::from_bits(b))],
                )
                .unwrap();
            assert_eq!(got, Value::Bool(oracle(a, b)), "x={a:016x} sign={b:016x}");
        }
    }
    output(
        "GATES",
        &serde_json::json!({"checks":cases.len()*2,"sizes":sizes}),
    );
    output(
        "CORPUS",
        &serde_json::json!({"pairs":cases.into_iter().map(|(a,b)|(format!("{a:016x}"),format!("{b:016x}"))).collect::<Vec<_>>()}),
    );
}
#[test]
fn other_carriers_keep_results_or_errors() {
    let vals = vec![
        Value::F64(f64::INFINITY),
        Value::F64(f64::NEG_INFINITY),
        Value::F64(f64::NAN),
        Value::F64(-0.),
        Value::F64(0.),
        Value::F64(1.),
        Value::I32(-1),
        Value::I32(0),
        Value::I64(i64::MAX),
        Value::Null,
        Value::Undefined,
        Value::Bool(true),
        Value::Bool(false),
        Value::String("1".into()),
        Value::String("bad".into()),
    ];
    let (mut a, fa) = worker(program(false, 0, 1));
    let (mut b, fb) = worker(program(true, 0, 1));
    let mut checks = 0;
    for x in &vals {
        for s in &vals {
            let old = a
                .invoke(&fa, &[x.clone(), s.clone()])
                .map_err(|e| e.message);
            let new = b
                .invoke(&fb, &[x.clone(), s.clone()])
                .map_err(|e| e.message);
            assert_eq!(old, new, "x={x:?} s={s:?}");
            checks += 1;
        }
    }
    output("CARRIERS", &serde_json::json!({"pairs":checks}));
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, op, a, b, control) in [
        ("positive-infinity", 0, f64::INFINITY, 1., false),
        ("negative-infinity", 0, f64::NEG_INFINITY, -1., false),
        ("either-zero-sign", 0, f64::INFINITY, -0., false),
        ("finite", 0, 1234567., 1., false),
        ("nan-sign", 0, f64::INFINITY, f64::NAN, false),
        ("control", 0, 1234567., 1., true),
    ] {
        for repeats in [1, 16] {
            let (va, fa) = worker(program(false, op, repeats));
            let (vb, fb) = worker(program(!control, op, repeats));
            let mut vs = [va, vb];
            let fs = [fa, fb];
            let args = vec![Value::F64(a), Value::F64(b)];
            for side in 0..2 {
                assert_eq!(
                    vs[side].invoke(&fs[side], &args).unwrap(),
                    Value::Bool(oracle(a.to_bits(), b.to_bits()))
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
